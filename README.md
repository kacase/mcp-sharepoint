# SharePoint MCP (Model Context Protocol) Server

A Model Context Protocol server that integrates with Microsoft SharePoint through Microsoft Graph API, allowing Claude and other LLMs to explore SharePoint sites, browse documents, and search content.

## Features

- 🏢 **Site Discovery**: List and explore all accessible SharePoint sites
- 📁 **Content Browsing**: Navigate document libraries, folders, and files
- 🔍 **Search Integration**: Search across all SharePoint content with natural language queries
- 📄 **File Access**: Download and view file content (text files as UTF-8, binary files as base64)
- 📚 **Smart Resources**: Automatic content discovery through MCP resources
- 🛡️ **Type Safety**: Full TypeScript implementation with Zod validation
- 🔐 **Read-Only**: Secure read-only access following principle of least privilege

## Prerequisites

- Node.js 18+
- Microsoft 365 account with SharePoint access
- Microsoft Azure App Registration with Graph API permissions

## Setup

1. Register an application in Azure Active Directory:
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to "App registrations"
   - Create a new registration with a redirect URI of type "Public client/native (mobile & desktop)"
     - Register `http://localhost` as the redirect URI

   - Configure API permissions:
     - Choose Microsoft Graph and type delegated, as we will act on the user's behalf
     - Add "Sites.Read.All" - Read access to all SharePoint sites
     - Add "Files.Read.All" - Read access to all files
     - Add "User.Read" - Basic user profile (required)

2. Note the values from your Azure app registration (Overview) to use for the MCP config as environment variables:
  - Client ID (Application (client) ID)
  - Authority ID (Directory (tenant) ID)

3. Register the MCP server
For Claude Desktop, create or update your configuration in `~/.claude/config.json`:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "npx",
      "args": [
        "mcp-sharepoint"
      ],
      "env": {
        "AUTHORITY": "your-authority-id",
        "CLIENT_ID": "your-client-id",
        "MCP_SERVER_NAME": "sharepoint-mcp",
        "MCP_SERVER_VERSION": "1.0.0"
      }
    }
  }
}
```

Make sure to replace the environment variables with your actual values.

## Available Tools

### Authentication
- **refreshAuthToken**: Refreshes authentication token to get updated permissions

### Site Discovery
- **listSharePointSites**: Lists all accessible SharePoint sites (your starting point)
- **getSharePointSite**: Gets details of a specific site
- **getSharePointSubsites**: Lists subsites within a site
- **getRootSharePointSite**: Gets the organization's main SharePoint site

### Content Exploration
- **listSiteLists**: Lists all lists in a SharePoint site
- **getSiteList**: Gets details of a specific list
- **listSiteDrives**: Lists document libraries in a site
- **getSiteDefaultDrive**: Gets the main document library of a site

### File Operations
- **listDriveItems**: Lists files and folders (use path parameter to navigate)
- **getDriveItem**: Gets details of a specific file or folder
- **getDriveItemContent**: Downloads file content (text or base64 for binary)

### Search
- **searchSharePoint**: Search across all SharePoint content with natural language

## Resources (Automatic Context)

Resources provide automatic context to the LLM about available SharePoint content:

### Static Resources
- **sharepoint://sites/all**: Complete list of accessible sites
- **sharepoint://sites/root**: Organization's main SharePoint site

### Dynamic Resources (Use actual IDs)
- **sharepoint://sites/{siteId}/structure**: Site's lists and document libraries
- **sharepoint://sites/{siteId}/files/{path}**: Browse files in specific folder
- **sharepoint://search/{query}**: Search results for specific query

## Usage Flow

1. **Discover Sites**: The `sharepoint://sites/all` resource shows all available sites
2. **Explore Structure**: Use `sharepoint://sites/{siteId}/structure` to see site contents
3. **Browse Files**: Use `sharepoint://sites/{siteId}/files/root` for top-level files
4. **Get Content**: Use `getDriveItemContent` tool to download specific files
5. **Search**: Use `searchSharePoint` tool for finding specific content

## Development

Run in development mode with live reloading:
```bash
npm run dev
```

Build the project:
```bash
npm run build
```

Run linting:
```bash
npm run lint
```

## Local Development Configuration

For local development, use an absolute path in your MCP configuration:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": [
        "/ABSOLUTE/PATH/TO/sharepoint_mcp/build/index.js"
      ],
      "env": {
        "AUTHORITY": "your-authority-id",
        "CLIENT_ID": "your-client-id",
        "MCP_SERVER_NAME": "sharepoint-mcp",
        "MCP_SERVER_VERSION": "1.0.0"
      }
    }
  }
}
```

## Security & Permissions

This server uses minimal read-only permissions:
- **Sites.Read.All**: Access to SharePoint sites and structure
- **Files.Read.All**: Access to read files and folders
- **User.Read**: Basic user profile information

No write permissions are requested, ensuring safe read-only access to your SharePoint content.