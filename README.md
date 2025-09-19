# Lark Doc CLI

A thin wrapper around the [Lark/Feishu Node SDK](https://github.com/larksuite/node-sdk) that lets you perform simple create, read, update and delete (CRUD) operations on Feishu documents from the command line.

## Prerequisites

- Node.js 16 or above
- A Feishu application with the Docs permissions enabled. Set the credentials via environment variables or CLI flags.

Create a `.env` file (optional) and provide your credentials:

```
LARK_APP_ID=cli_a84736eb4038900d
LARK_APP_SECRET=LnZtpqFQBJGfQJlpbwu4RdprJSO2TYOB
# LARK_TENANT_KEY=your-tenant-key     # only needed for ISV apps
# LARK_USER_ACCESS_TOKEN=xxx          # required for user scoped operations
# LARK_DEFAULT_COLLABORATORS=openid:ou_xxx:edit,email:user@example.com:view
```

Then install dependencies:

```
npm install
```

## Usage

Run the CLI with `npm start -- <command>` or install it globally with `npm link`.

### Create a document

```
npm start -- create --title "Demo" --content "Hello from CLI"
```

Automatically add collaborators by repeating `--collaborator memberType:memberId[:perm]`:

```
npm start -- create --title "Demo" --collaborator openid:ou_xxx:edit --collaborator email:user@example.com:view
```

You can also seed the document from a file:

```
npm start -- create --title "Demo" --content-file notes.txt
```

### Read document metadata (and optional raw content)

```
npm start -- read <documentToken> --raw
```

### Update document content

```
npm start -- update <documentToken> --content "Updated content"
```

### Delete a document

```
npm start -- delete <documentToken>
```

Use `--help` on any command to see the available flags. You can override credentials per command:

```
npm start -- create --title Test --app-id xxx --app-secret yyy
```

## Notes

- The CLI replaces the entire document body during `update`. Multi-line strings are split on newlines and each line becomes a paragraph block.
- Some APIs require a user access token. Provide it with `--user-access-token` (or `LARK_USER_ACCESS_TOKEN`) if the tenant-scoped credential is not sufficient.
- Deleting a document moves it to the recycle bin via the Drive API; permanent deletion must be done from Feishu.
- Set `LARK_DEFAULT_COLLABORATORS` (comma separated `memberType:memberId[:perm]`) to share every created doc automatically. Supported `memberType` values include `openid`, `userid`, `email`, `unionid`, etc. Permissions default to `edit`.
