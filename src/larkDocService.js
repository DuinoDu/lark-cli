const lark = require('@larksuiteoapi/node-sdk');

const { Domain, withAll, withTenantKey, withUserAccessToken } = lark;

function buildRequestOptions({ tenantKey, userAccessToken }) {
  const options = [];
  if (tenantKey) {
    options.push(withTenantKey(tenantKey));
  }
  if (userAccessToken) {
    options.push(withUserAccessToken(userAccessToken));
  }
  if (!options.length) {
    return undefined;
  }
  return options.length === 1 ? options[0] : withAll(options);
}

function toParagraphBlock(line) {
  return {
    block_type: 1,
    text: {
      elements: [
        {
          text_run: {
            content: line,
          },
        },
      ],
    },
  };
}

class LarkDocService {
  constructor({ appId, appSecret, domain, tenantKey, userAccessToken } = {}) {
    if (!appId || !appSecret) {
      throw new Error('LARK_APP_ID and LARK_APP_SECRET must be provided');
    }

    this.client = new lark.Client({
      appId,
      appSecret,
      domain: domain || Domain.Feishu,
    });

    this.requestOptions = buildRequestOptions({ tenantKey, userAccessToken });
  }

  async createDocument({ title, folderToken, content, collaborators }) {
    const response = await this.client.docx.document.create(
      {
        data: {
          title,
          folder_token: folderToken,
        },
      },
      this.requestOptions,
    );

    const document = response?.data?.document;
    if (!document?.document_id) {
      throw new Error('Unable to create document. Response missing document id');
    }

    if (content !== undefined) {
      await this.replaceDocumentContent(document.document_id, content);
    }

    if (Array.isArray(collaborators) && collaborators.length > 0) {
      await this.addCollaborators(document.document_id, collaborators);
    }

    return document;
  }

  async getDocument(documentId) {
    const response = await this.client.docx.document.get(
      {
        path: { document_id: documentId },
      },
      this.requestOptions,
    );
    return response?.data?.document;
  }

  async getRawContent(documentId) {
    const response = await this.client.docx.document.rawContent(
      {
        path: { document_id: documentId },
      },
      this.requestOptions,
    );
    return response?.data?.content ?? '';
  }

  async deleteDocument(documentId) {
    await this.client.drive.file.delete(
      {
        path: { file_token: documentId },
      },
      this.requestOptions,
    );
  }

  async replaceDocumentContent(documentId, rawText) {
    const pageBlock = await this._getPageBlock(documentId);
    const childrenCount = pageBlock?.children?.length || 0;
    const path = { document_id: documentId, block_id: pageBlock.block_id };

    if (childrenCount > 0) {
      await this.client.docx.documentBlockChildren.batchDelete(
        {
          path,
          data: {
            start_index: 0,
            end_index: childrenCount,
          },
        },
        this.requestOptions,
      );
    }

    const lines = rawText.split(/\r?\n/);
    if (!lines.length) {
      return;
    }

    const children = lines.map(toParagraphBlock);
    await this.client.docx.documentBlockChildren.create(
      {
        path,
        data: {
          index: 0,
          children,
        },
      },
      this.requestOptions,
    );
  }

  async addCollaborators(documentId, collaborators) {
    if (!Array.isArray(collaborators) || collaborators.length === 0) {
      return;
    }

    const normalized = collaborators
      .map((item) => {
        if (!item) {
          return null;
        }
        const {
          memberType,
          memberId,
          perm = 'edit',
          needNotification = false,
          permType,
          entityType,
        } = item;

        if (!memberType || !memberId) {
          return null;
        }

        return {
          member_type: memberType,
          member_id: memberId,
          perm,
          need_notification: needNotification,
          perm_type: "container",
          type: "user",
        };
      })
      .filter(Boolean);

    if (!normalized.length) {
      return;
    }

    for (const collaborator of normalized) {
      try {
        const { need_notification, ...data } = collaborator;
        await this.client.drive.permission.member.create(
          {
            path: { token: documentId },
            params: {
              type: 'docx',
              need_notification: need_notification,
            },
            data,
          },
          this.requestOptions,
        );
      } catch (error) {
        console.error('Failed to add collaborator:', collaborator, error.message);
        // Continue with other collaborators
      }
    }
  }

  async _getPageBlock(documentId) {
    const response = await this.client.docx.documentBlock.list(
      {
        path: { document_id: documentId },
        params: {
          page_size: 500,
        },
      },
      this.requestOptions,
    );

    const items = response?.data?.items || [];
    const pageBlock = items.find((item) => item?.page);
    if (!pageBlock) {
      throw new Error('Unable to locate page block for document');
    }
    return pageBlock;
  }
}

module.exports = {
  LarkDocService,
};
