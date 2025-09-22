#!/usr/bin/env node

require('dotenv').config({ quiet: true });

const fs = require('fs');
const path = require('path');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const { LarkDocService } = require('./larkDocService');

function resolveContentInput({ content, contentFile }) {
  if (content !== undefined && contentFile) {
    throw new Error('Use either --content or --content-file, not both');
  }
  if (contentFile) {
    const fullPath = path.resolve(contentFile);
    return fs.readFileSync(fullPath, 'utf8');
  }
  if (content !== undefined) {
    return content;
  }
  return undefined;
}

function parseCollaboratorEntry(entry) {
  const value = String(entry || '').trim();
  if (!value) {
    return null;
  }

  const parts = value.split(':');
  if (parts.length < 2) {
    throw new Error(`Invalid collaborator definition: "${value}". Use memberType:memberId[:perm]`);
  }

  const [memberType, memberId, perm] = parts;
  return {
    memberType,
    memberId,
    perm: perm || 'edit',
  };
}

function createService(argv) {
  const appId = argv.appId || process.env.LARK_APP_ID;
  const appSecret = argv.appSecret || process.env.LARK_APP_SECRET;
  const tenantKey = argv.tenantKey || process.env.LARK_TENANT_KEY;
  const userAccessToken = argv.userAccessToken || process.env.LARK_USER_ACCESS_TOKEN;

  return new LarkDocService({
    appId,
    appSecret,
    tenantKey,
    userAccessToken,
  });
}

async function handleCreate(argv) {
  const service = createService(argv);
  const content = resolveContentInput(argv);

  // Get wiki configuration from environment or command line
  const moveToWiki = argv.wiki || process.env.LARK_WIKI_AUTO_MOVE === 'true';
  const wikiSpaceId = argv.wikiSpace || process.env.LARK_WIKI_SPACE_ID;
  const wikiNodeId = argv.wikiNode || process.env.LARK_WIKI_ROOT_ID;

  const document = await service.createDocument({
    title: argv.title,
    folderToken: argv.folder,
    content,
    moveToWiki,
    wikiSpaceId,
    wikiNodeId,
  });

  const output = {
    document_id: document.document_id,
    title: document.title,
    revision_id: document.revision_id,
  };

  // Add wiki info if document was moved
  if (document.moved_to_wiki) {
    output.wiki_node_token = document.wiki_node_token;
    output.moved_to_wiki = true;
  }

  console.log(JSON.stringify(output, null, 2));
}

async function handleRead(argv) {
  const service = createService(argv);
  const document = await service.getDocument(argv.documentId);
  if (!document) {
    console.error('Document not found');
    process.exit(1);
  }
  const result = { document };
  if (argv.raw || argv.showContent) {
    result.raw_content = await service.getRawContent(argv.documentId);
  }
  console.log(JSON.stringify(result, null, 2));
}

async function handleUpdate(argv) {
  const service = createService(argv);
  const content = resolveContentInput(argv);
  if (content === undefined) {
    throw new Error('Update command requires --content or --content-file');
  }
  await service.appendDocumentContent(argv.documentId, content, argv.markdown);
  console.log(`Document ${argv.documentId} updated`);
}

async function handleDelete(argv) {
  const service = createService(argv);
  await service.deleteDocument(argv.documentId);
  console.log(`Document ${argv.documentId} deleted`);
}

yargs(hideBin(process.argv))
  .scriptName('lark-doc')
  .usage('$0 <command> [options]')
  .option('app-id', {
    describe: 'Feishu app id',
    type: 'string',
    default: process.env.LARK_APP_ID,
  })
  .option('app-secret', {
    describe: 'Feishu app secret',
    type: 'string',
    default: process.env.LARK_APP_SECRET,
  })
  .option('tenant-key', {
    describe: 'Tenant key for ISV apps',
    type: 'string',
    default: process.env.LARK_TENANT_KEY,
  })
  .option('user-access-token', {
    describe: 'User access token when required by the API scope',
    type: 'string',
    default: process.env.LARK_USER_ACCESS_TOKEN,
  })
  .command(
    'create',
    'Create a new Feishu doc',
    (cmd) =>
      cmd
        .option('title', {
          alias: 't',
          describe: 'Document title',
          type: 'string',
          demandOption: true,
        })
        .option('folder', {
          alias: 'f',
          describe: 'Folder token to place the doc under',
          type: 'string',
        })
        .option('content', {
          alias: 'c',
          describe: 'Initial content for the doc',
          type: 'string',
        })
        .option('content-file', {
          describe: 'Path to a file used as the initial content',
          type: 'string',
        })
        .option('collaborator', {
          alias: 'a',
          describe: 'Default collaborator definition in the form memberType:memberId[:perm] (repeatable)',
          type: 'string',
          array: true,
        })
        .option('markdown', {
          alias: 'm',
          describe: 'Treat content as markdown and convert to document blocks',
          type: 'boolean',
          default: false,
        })
        .option('wiki', {
          alias: 'w',
          describe: 'Move document to wiki after creation',
          type: 'boolean',
          default: false,
        })
        .option('wiki-space', {
          describe: 'Wiki space ID (overrides LARK_WIKI_SPACE_ID env var)',
          type: 'string',
        })
        .option('wiki-node', {
          describe: 'Wiki node ID (overrides LARK_WIKI_ROOT_ID env var)',
          type: 'string',
        }),
    (argv) =>
      handleCreate(argv).catch((err) => {
        console.error(err.message || err);
        process.exit(1);
      }),
  )
  .command(
    'read <documentId>',
    'Fetch document metadata and optional content',
    (cmd) =>
      cmd
        .positional('documentId', {
          describe: 'Document token/id',
          type: 'string',
        })
        .option('raw', {
          alias: 'r',
          describe: 'Include raw text content',
          type: 'boolean',
          default: false,
        })
        .option('show-content', {
          describe: 'Alias for --raw',
          type: 'boolean',
          default: false,
        }),
    (argv) =>
      handleRead(argv).catch((err) => {
        console.error(err.message || err);
        process.exit(1);
      }),
  )
  .command(
    'update <documentId>',
    'Append content to the document',
    (cmd) =>
      cmd
        .positional('documentId', {
          describe: 'Document token/id',
          type: 'string',
        })
        .option('content', {
          alias: 'c',
          describe: 'Content to append to the document',
          type: 'string',
        })
        .option('content-file', {
          describe: 'Path to file that contains content to append',
          type: 'string',
        })
        .option('markdown', {
          alias: 'm',
          describe: 'Treat content as markdown and convert to document blocks',
          type: 'boolean',
          default: false,
        }),
    (argv) =>
      handleUpdate(argv).catch((err) => {
        console.error(err.message || err);
        process.exit(1);
      }),
  )
  .command(
    'delete <documentId>',
    'Move the document to the recycle bin',
    (cmd) =>
      cmd.positional('documentId', {
        describe: 'Document token/id',
        type: 'string',
      }),
    (argv) =>
      handleDelete(argv).catch((err) => {
        console.error(err.message || err);
        process.exit(1);
      }),
  )
  .demandCommand(1, 'Please specify a command')
  .strict()
  .help()
  .wrap(Math.min(100, yargs().terminalWidth()))
  .parse();
