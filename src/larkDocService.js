const lark = require('@larksuiteoapi/node-sdk');

const { Domain, withAll, withTenantKey, withUserAccessToken } = lark;

function _logApiCall(description, functionName, params = {}) {
  console.log(`[Lark API] ${description} - ${functionName}: ${JSON.stringify(params, null, 2)}`);
}

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
  const MAX_CHUNK_SIZE = 10000; // Maximum 100,000 UTF-16 characters

  // If the line is within the limit, create a single block
  if (line.length <= MAX_CHUNK_SIZE) {
    return {
      block_type: 2,
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

  // For lines exceeding the limit, split into multiple blocks
  const blocks = [];
  let start = 0;

  while (start < line.length) {
    const end = Math.min(start + MAX_CHUNK_SIZE, line.length);
    const chunk = line.substring(start, end);

    blocks.push({
      block_type: 2,
      text: {
        elements: [
          {
            text_run: {
              content: chunk,
            },
          },
        ],
      },
    });

    start = end;
  }

  return blocks;
}

function processTextToBlocks(text) {
  const lines = text.split(/\r?\n/);
  const blocks = [];

  for (const line of lines) {
    const result = toParagraphBlock(line);
    if (Array.isArray(result)) {
      blocks.push(...result);
    } else {
      blocks.push(result);
    }
  }

  return blocks;
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

  async createDocument({ title, folderToken, content, moveToWiki = false, wikiSpaceId, wikiNodeId, markdown = false }) {
    _logApiCall('Create document', 'docx.document.create', { title, folderToken });
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
      await this._insertContentToDocument(document.document_id, content, markdown);
    }

    // Move document to wiki if requested
    if (moveToWiki && wikiSpaceId && wikiNodeId) {
      try {
        const wikiNode = await this.moveDocumentToWiki(document.document_id, wikiSpaceId, wikiNodeId);
        if (wikiNode) {
          // Delete the original document after successful move
          await this.deleteDocument(document.document_id);
          return {
            ...document,
            wiki_node_token: wikiNode.node_token,
            moved_to_wiki: true
          };
        }
      } catch (error) {
        console.error('Failed to move document to wiki, keeping original document:', error.message);
        // Keep the original document if move fails
      }
    }

    return document;
  }

  async getDocument(documentId) {
    _logApiCall('Get document metadata', 'docx.document.get', { documentId });
    const response = await this.client.docx.document.get(
      {
        path: { document_id: documentId },
      },
      this.requestOptions,
    );
    return response?.data?.document;
  }

  async getRawContent(documentId) {
    _logApiCall('Get document raw content', 'docx.document.rawContent', { documentId });
    const response = await this.client.docx.document.rawContent(
      {
        path: { document_id: documentId },
      },
      this.requestOptions,
    );
    return response?.data?.content ?? '';
  }

  async deleteDocument(documentId) {
    _logApiCall('Delete document', 'drive.file.delete', { documentId });
    await this.client.drive.file.delete(
      {
        path: { file_token: documentId },
      },
      this.requestOptions,
    );
  }

  async appendDocumentContent(documentId, rawText, isMarkdown = false) {
    if (!rawText) {
      return;
    }

    // Append content to the end of the document
    await this._appendContentToDocument(documentId, rawText, isMarkdown);
  }

  async _appendContentToDocument(documentId, content, isMarkdown = false) {
    // Step 1: Convert content to document blocks
    let blocks;
    if (isMarkdown) {
      blocks = await this._convertContentToBlocks(content);
    } else {
      // For plain text, create simple paragraph blocks
      blocks = processTextToBlocks(content);
    }

    // Step 2: Get document blocks to find the page block
    const documentBlocks = await this._getDocumentBlocks(documentId);

    // Step 3: Append content at the end of the document
    if (documentBlocks.length > 0) {
      const pageBlock = documentBlocks.find((block) => block?.page);
      if (pageBlock) {
        const childrenCount = pageBlock?.children?.length || 0;
        _logApiCall('Append content blocks to document', 'docx.documentBlockChildren.create', {
          documentId,
          blockId: pageBlock.block_id,
          index: childrenCount,
          blockCount: blocks.length
        });
        await this.client.docx.documentBlockChildren.create(
          {
            path: { document_id: documentId, block_id: pageBlock.block_id },
            data: {
              index: childrenCount,
              children: blocks,
            },
          },
          this.requestOptions,
        );
      }
    }
  }

  async _insertContentToDocument(documentId, content, isMarkdown = false) {
    // Step 1: Convert content to document blocks
    let blocks;

    if (isMarkdown) {
      blocks = await this._convertContentToBlocks(content);
    } else {
      // For plain text, create simple paragraph blocks
      blocks = processTextToBlocks(content);
    }

    // Step 2: Get document blocks to find the page block and count children
    const documentBlocks = await this._getDocumentBlocks(documentId);

    // Step 3: Insert content at the end of the document
    if (documentBlocks.length > 0) {
      const pageBlock = documentBlocks.find((block) => block?.page);
      if (pageBlock) {
        const childrenCount = pageBlock?.children?.length || 0;

        // If blocks.length exceeds 50, split into multiple API calls
        const MAX_BLOCKS_PER_CALL = 50;

        if (blocks.length <= MAX_BLOCKS_PER_CALL) {
          _logApiCall('Insert content blocks to document', 'docx.documentBlockChildren.create', {
            documentId,
            blockId: pageBlock.block_id,
            index: childrenCount,
            blockCount: blocks.length
          });
          await this.client.docx.documentBlockChildren.create(
            {
              path: { document_id: documentId, block_id: pageBlock.block_id },
              data: {
                index: childrenCount,
                children: blocks,
              },
            },
            this.requestOptions,
          );
        } else {
          // Split blocks into chunks of MAX_BLOCKS_PER_CALL or less
          let currentIndex = childrenCount;
          for (let i = 0; i < blocks.length; i += MAX_BLOCKS_PER_CALL) {
            const chunk = blocks.slice(i, i + MAX_BLOCKS_PER_CALL);
            _logApiCall('Insert content blocks to document (chunked)', 'docx.documentBlockChildren.create', {
              documentId,
              blockId: pageBlock.block_id,
              index: currentIndex,
              chunk: i / MAX_BLOCKS_PER_CALL + 1,
              totalChunks: Math.ceil(blocks.length / MAX_BLOCKS_PER_CALL),
              blockCount: chunk.length
            });
            await this.client.docx.documentBlockChildren.create(
              {
                path: { document_id: documentId, block_id: pageBlock.block_id },
                data: {
                  index: currentIndex,
                  children: chunk,
                },
              },
              this.requestOptions,
            );
            currentIndex += chunk.length;
          }
        }
      }
    }
  }

  async _convertContentToBlocks(content) {
    try {
      _logApiCall('Convert content to document blocks', 'docx.document.convert', {
        contentLength: content.length,
        contentType: 'markdown'
      });
      const response = await this.client.docx.document.convert(
        {
          data: {
            content,
            content_type: 'markdown',
          },
        },
        this.requestOptions,
      );
      const normalizedBlocks = this._reshapeConvertedBlocks(response?.data);
      if (normalizedBlocks.length) {
        return normalizedBlocks;
      }
      return processTextToBlocks(content);
    } catch (error) {
      console.error('Failed to convert content to blocks, falling back to simple text:', error.message);
      // Fallback to simple text processing
      return processTextToBlocks(content);
    }
  }

  _reshapeConvertedBlocks(convertedData = {}) {
    const { blocks, first_level_block_ids: firstLevelBlockIds } = convertedData;
    if (!blocks || !firstLevelBlockIds?.length) {
      return [];
    }

    const blockMap = Array.isArray(blocks)
      ? blocks.reduce((acc, block) => {
        if (block?.block_id) {
          acc[block.block_id] = block;
        }
        return acc;
      }, {})
      : blocks;

    const buildBlock = (blockId) => {
      const block = blockMap?.[blockId];
      if (!block) {
        return null;
      }

      const { block_id: _ignoredBlockId, parent_id: _ignoredParentId, children, ...rest } = block;
      const sanitizedBlock = { ...rest };

      if (Array.isArray(children) && children.length) {
        const childBlocks = children
          .map((childId) => buildBlock(childId))
          .filter(Boolean);
        if (childBlocks.length) {
          sanitizedBlock.children = childBlocks;
        }
      }

      return sanitizedBlock;
    };

    return firstLevelBlockIds
      .map((blockId) => buildBlock(blockId))
      .filter(Boolean);
  }

  async _getDocumentBlocks(documentId) {
    _logApiCall('Get document blocks', 'docx.documentBlock.list', { documentId, pageSize: 500 });
    const response = await this.client.docx.documentBlock.list(
      {
        path: { document_id: documentId },
        params: {
          page_size: 500,
        },
      },
      this.requestOptions,
    );
    return response?.data?.items || [];
  }

  async _getPageBlock(documentId) {
    const items = await this._getDocumentBlocks(documentId);
    const pageBlock = items.find((item) => item?.page);
    if (!pageBlock) {
      throw new Error('Unable to locate page block for document');
    }
    return pageBlock;
  }

  async moveDocumentToWiki(documentId, spaceId, nodeId) {
    try {
      _logApiCall('Move document to wiki', 'wiki.spaceNode.moveDocsToWiki', {
        documentId,
        spaceId,
        nodeId,
        objType: 'docx'
      });
      const response = await this.client.wiki.spaceNode.moveDocsToWiki(
        {
          path: { space_id: spaceId },
          data: {
            parent_wiki_token: nodeId,
            obj_type: 'docx',
            obj_token: documentId,
          },
        },
        this.requestOptions,
      );

      return response?.data?.node;
    } catch (error) {
      console.error('Failed to move document to wiki:', error.message);
      throw new Error(`Failed to move document to wiki: ${error.message}`);
    }
  }
}

module.exports = {
  LarkDocService,
};
