#!/usr/bin/env node

require('dotenv').config({ quiet: true });

const fs = require('fs');
const path = require('path');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const { LarkDocService } = require('./larkDocService');

function extractDocumentIdFromUrl(input) {
  if (!input) {
    return null;
  }

  const trimmed = String(input).trim();
  if (!trimmed) {
    return null;
  }

  // If the input already looks like a token (no schema, no slashes), accept it as-is
  if (!trimmed.includes('://') && !trimmed.includes('/')) {
    return trimmed;
  }

  try {
    const url = new URL(trimmed);
    const segments = url.pathname.split('/').filter(Boolean);
    if (!segments.length) {
      return null;
    }

    const docTypeIndex = segments.findIndex((segment) => segment === 'docx' || segment === 'docs');
    if (docTypeIndex >= 0 && segments[docTypeIndex + 1]) {
      return segments[docTypeIndex + 1];
    }

    // Fallback to last segment when doc type is not explicitly present
    return segments[segments.length - 1];
  } catch (error) {
    return null;
  }
}

function sanitizeFileName(name) {
  const replaced = String(name || '')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!replaced) {
    return 'document';
  }

  // Limit to a reasonable length to avoid filesystem issues
  return replaced.slice(0, 120);
}

function resolveOutputPath({ output, title }) {
  const sanitizedTitle = sanitizeFileName(title);
  const defaultFileName = `${sanitizedTitle}.md`;

  if (!output) {
    return path.resolve(defaultFileName);
  }

  const resolved = path.resolve(output);
  if (fs.existsSync(resolved) && fs.statSync(resolved).isDirectory()) {
    return path.join(resolved, defaultFileName);
  }

  if (path.extname(resolved).toLowerCase() === '.md') {
    return resolved;
  }

  return `${resolved}.md`;
}

function resolveContentInput({ content, contentFile }) {
  if (content !== undefined && contentFile) {
    throw new Error('Use either --content or --content-file, not both');
  }
  if (contentFile) {
    const fullPath = path.resolve(contentFile);
    const value = fs.readFileSync(fullPath, 'utf8');
    const isMarkdownHint = path.extname(fullPath).toLowerCase() === '.md';
    const titleHint = path.basename(fullPath, path.extname(fullPath));
    return { content: value, isMarkdownHint, titleHint };
  }
  if (content !== undefined) {
    return { content, isMarkdownHint: false, titleHint: undefined };
  }
  return { content: undefined, isMarkdownHint: false, titleHint: undefined };
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
  const { content, isMarkdownHint, titleHint } = resolveContentInput(argv);
  const title = argv.title || titleHint;
  if (!title) {
    throw new Error('Document title is required. Provide --title or use --content-file to derive it from the file name.');
  }

  // Get wiki configuration from environment or command line
  const moveToWiki = argv.wiki || process.env.LARK_WIKI_AUTO_MOVE === 'true';
  const wikiSpaceId = argv.wikiSpace || process.env.LARK_WIKI_SPACE_ID;
  const wikiNodeId = argv.wikiNode || process.env.LARK_WIKI_ROOT_ID;

  const document = await service.createDocument({
    title,
    folderToken: argv.folder,
    content,
    markdown: argv.markdown || isMarkdownHint,
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
  const { content, isMarkdownHint } = resolveContentInput(argv);
  if (content === undefined) {
    throw new Error('Update command requires --content or --content-file');
  }
  await service.appendDocumentContent(argv.documentId, content, argv.markdown || isMarkdownHint);
  console.log(`Document ${argv.documentId} updated`);
}

async function handleDelete(argv) {
  const service = createService(argv);
  await service.deleteDocument(argv.documentId);
  console.log(`Document ${argv.documentId} deleted`);
}

async function handleDownload(argv) {
  const documentId = extractDocumentIdFromUrl(argv.docUrl);
  if (!documentId) {
    throw new Error('Unable to determine document token from the provided URL');
  }

  const service = createService(argv);
  const document = await service.getDocument(documentId);
  if (!document) {
    throw new Error(`Document ${documentId} not found`);
  }

  const rawContent = await service.getRawContent(documentId);
  const outputPath = resolveOutputPath({ output: argv.output, title: document.title || documentId });

  if (!argv.force && fs.existsSync(outputPath)) {
    throw new Error(`File already exists at ${outputPath}. Use --force to overwrite.`);
  }

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, rawContent ?? '', 'utf8');

  console.log(`Markdown saved to ${outputPath}`);
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
          describe: 'Document title (defaults to content file name when omitted)',
          type: 'string',
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
  .command(
    'download <docUrl>',
    'Download a Feishu doc as local markdown',
    (cmd) =>
      cmd
        .positional('docUrl', {
          describe: 'Feishu doc URL or document token',
          type: 'string',
        })
        .option('output', {
          alias: 'o',
          describe: 'Output file path (defaults to document title). Directories will contain <title>.md',
          type: 'string',
        })
        .option('force', {
          alias: 'f',
          describe: 'Overwrite the output file if it already exists',
          type: 'boolean',
          default: false,
        }),
    (argv) =>
      handleDownload(argv).catch((err) => {
        console.error(err.message || err);
        process.exit(1);
      }),
  )
  .demandCommand(1, 'Please specify a command')
  .strict()
  .help()
  .wrap(Math.min(100, yargs().terminalWidth()))
  .parse();
