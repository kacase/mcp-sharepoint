#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { 
  ListSitesQuerySchema,
  ListDriveItemsQuerySchema,
  GetDriveItemContentSchema,
  SearchSharePointSchema
} from "./types.js";
import { graphClient } from "./graphClient.js";
import { refreshToken } from "./auth.js";

const server = new McpServer({
  name: "sharepoint-mcp",
  version: "1.0.0"
});

// ============= Authentication Tools =============

server.tool(
  "refreshAuthToken",
  "Refreshes the authentication token to get updated SharePoint permissions. Use this if you get permission errors after updating your Azure AD app registration.",
  {},
  async () => {
    try {
      await refreshToken();
      
      return {
        content: [
          {
            type: "text",
            text: "Authentication token successfully refreshed with updated permissions. You can now access SharePoint with the new scopes."
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error refreshing authentication token: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= SharePoint Sites Tools =============

server.tool(
  "listSharePointSites",
  "Lists all SharePoint sites accessible to the user. Use optional 'search' parameter to filter by name, or 'top' to limit results. This is your starting point for discovering what SharePoint content is available.",
  ListSitesQuerySchema.shape,
  async (params) => {
    try {
      const sites = await graphClient.listSites(params);
      
      const formattedSites = sites.map(site => ({
        id: site.id,
        name: site.displayName || site.name || 'Unnamed Site',
        description: site.description || '',
        webUrl: site.webUrl,
        createdDateTime: site.createdDateTime ? new Date(site.createdDateTime).toLocaleString() : 'Unknown',
        lastModifiedDateTime: site.lastModifiedDateTime ? new Date(site.lastModifiedDateTime).toLocaleString() : 'Unknown',
        hostname: site.siteCollection?.hostname || ''
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedSites, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing SharePoint sites: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getSharePointSite",
  "Gets details of a specific SharePoint site",
  {
    siteId: z.string().describe("ID of the SharePoint site to retrieve")
  },
  async (params) => {
    try {
      const site = await graphClient.getSite(params.siteId);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(site, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting SharePoint site: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getSharePointSubsites",
  "Gets subsites of a SharePoint site",
  {
    siteId: z.string().describe("ID of the parent SharePoint site")
  },
  async (params) => {
    try {
      const subsites = await graphClient.getSubsites(params.siteId);
      
      const formattedSubsites = subsites.map(site => ({
        id: site.id,
        name: site.displayName || site.name || 'Unnamed Site',
        description: site.description || '',
        webUrl: site.webUrl,
        createdDateTime: site.createdDateTime ? new Date(site.createdDateTime).toLocaleString() : 'Unknown'
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedSubsites, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting SharePoint subsites: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getRootSharePointSite",
  "Gets the organization's root SharePoint site",
  {},
  async () => {
    try {
      const rootSite = await graphClient.getRootSite();
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(rootSite, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting root SharePoint site: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= SharePoint Lists Tools =============

server.tool(
  "listSiteLists",
  "Lists all lists in a SharePoint site",
  {
    siteId: z.string().describe("ID of the SharePoint site")
  },
  async (params) => {
    try {
      const lists = await graphClient.listSiteLists(params.siteId);
      
      const formattedLists = lists.map(list => ({
        id: list.id,
        name: list.displayName || list.name || 'Unnamed List',
        description: list.description || '',
        webUrl: list.webUrl || '',
        createdDateTime: list.createdDateTime ? new Date(list.createdDateTime).toLocaleString() : 'Unknown',
        template: list.list?.template || 'Unknown',
        hidden: list.list?.hidden || false
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedLists, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing site lists: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getSiteList",
  "Gets details of a specific list in a SharePoint site",
  {
    siteId: z.string().describe("ID of the SharePoint site"),
    listId: z.string().describe("ID of the list to retrieve")
  },
  async (params) => {
    try {
      const list = await graphClient.getSiteList(params.siteId, params.listId);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(list, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting site list: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= SharePoint Drives (Document Libraries) Tools =============

server.tool(
  "listSiteDrives",
  "Lists all drives (document libraries) in a SharePoint site",
  {
    siteId: z.string().describe("ID of the SharePoint site")
  },
  async (params) => {
    try {
      const drives = await graphClient.listSiteDrives(params.siteId);
      
      const formattedDrives = drives.map(drive => ({
        id: drive.id,
        name: drive.name || 'Unnamed Drive',
        description: drive.description || '',
        webUrl: drive.webUrl || '',
        driveType: drive.driveType || 'Unknown',
        owner: drive.owner?.user?.displayName || 'Unknown',
        createdDateTime: drive.createdDateTime ? new Date(drive.createdDateTime).toLocaleString() : 'Unknown'
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedDrives, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing site drives: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getSiteDefaultDrive",
  "Gets the default drive (document library) of a SharePoint site",
  {
    siteId: z.string().describe("ID of the SharePoint site")
  },
  async (params) => {
    try {
      const drive = await graphClient.getSiteDefaultDrive(params.siteId);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(drive, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting site default drive: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= SharePoint Files and Folders Tools =============

server.tool(
  "listDriveItems",
  "Lists files and folders in a SharePoint drive or folder. Requires siteId. Use 'path' parameter to navigate to specific folders (e.g., 'Documents/Projects'). Returns file metadata including IDs needed for downloading content.",
  ListDriveItemsQuerySchema.shape,
  async (params) => {
    try {
      const items = await graphClient.listDriveItems(params);
      
      const formattedItems = items.map(item => ({
        id: item.id,
        name: item.name,
        webUrl: item.webUrl || '',
        size: item.size || 0,
        isFolder: !!item.folder,
        isFile: !!item.file,
        mimeType: item.file?.mimeType || null,
        childCount: item.folder?.childCount || null,
        createdDateTime: item.createdDateTime ? new Date(item.createdDateTime).toLocaleString() : 'Unknown',
        lastModifiedDateTime: item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleString() : 'Unknown',
        createdBy: item.createdBy?.user?.displayName || 'Unknown',
        lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Unknown'
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedItems, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing drive items: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getDriveItem",
  "Gets details of a specific file or folder in a SharePoint drive",
  {
    siteId: z.string().describe("ID of the SharePoint site"),
    driveId: z.string().optional().describe("ID of the drive (optional, uses default drive if not specified)"),
    itemId: z.string().describe("ID of the drive item to retrieve")
  },
  async (params) => {
    try {
      const item = await graphClient.getDriveItem(params.siteId, params.driveId, params.itemId);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(item, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting drive item: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getDriveItemContent",
  "Downloads the actual content of a file from SharePoint. Returns text files as UTF-8 text, binary files as base64. Use this after finding the file with listDriveItems. Requires siteId and itemId.",
  GetDriveItemContentSchema.shape,
  async (params) => {
    try {
      const result = await graphClient.getDriveItemContent(params);
      
      const responseText = result.isBase64 
        ? `Binary file content (base64 encoded, ${Math.round(result.content.length * 0.75)} bytes, MIME: ${result.mimeType})\n\n${result.content}`
        : `Text file content (MIME: ${result.mimeType}):\n\n${result.content}`;

      return {
        content: [
          {
            type: "text",
            text: responseText
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting drive item content: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= SharePoint Search Tools =============

server.tool(
  "searchSharePoint",
  "Searches across all SharePoint content including files, documents, lists, and sites. Use natural language queries (e.g., 'budget 2024', 'project proposal'). Optional: limit to specific site with 'siteId', filter by content types with 'entityTypes'.",
  SearchSharePointSchema.shape,
  async (params) => {
    try {
      const results = await graphClient.searchSharePoint(params);
      
      const formattedResults = results.map(result => ({
        id: result.id,
        name: result.name || result.title || 'Unnamed',
        webUrl: result.webUrl || '',
        summary: result.summary || '',
        hitHighlightedSummary: result.hitHighlightedSummary || '',
        resourceType: result.resource?.["@odata.type"] || 'Unknown'
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedResults, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error searching SharePoint: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= Resources =============

// Static resource: All accessible SharePoint sites
server.resource(
  "sharepoint://sites/all",
  "Complete list of all SharePoint sites you have access to. This provides an overview of available sites with their IDs, names, descriptions, and URLs. Use the site IDs from this list with other tools and resources.",
  async (uri) => {
    try {
      const sites = await graphClient.listSites({});
      
      const formattedSites = sites.map(site => ({
        id: site.id,
        name: site.displayName || site.name || 'Unnamed Site',
        description: site.description || '',
        webUrl: site.webUrl,
        hostname: site.siteCollection?.hostname || '',
        createdDateTime: site.createdDateTime ? new Date(site.createdDateTime).toLocaleString() : 'Unknown'
      }));

      return {
        contents: [
          {
            uri: uri.toString(),
            text: JSON.stringify(formattedSites, null, 2),
            mimeType: "application/json"
          }
        ]
      };
    } catch (error) {
      return {
        contents: [
          {
            uri: uri.toString(),
            text: `Error accessing SharePoint sites: ${error instanceof Error ? error.message : String(error)}`,
            mimeType: "text/plain"
          }
        ]
      };
    }
  }
);

// Static resource: Root/organization SharePoint site  
server.resource(
  "sharepoint://sites/root",
  "Your organization's main SharePoint site (root site). This is typically the central hub for company-wide content and resources.",
  async (uri) => {
    try {
      const rootSite = await graphClient.getRootSite();
      
      const formattedSite = {
        id: rootSite.id,
        name: rootSite.displayName || rootSite.name || 'Root Site',
        description: rootSite.description || '',
        webUrl: rootSite.webUrl,
        hostname: rootSite.siteCollection?.hostname || '',
        createdDateTime: rootSite.createdDateTime ? new Date(rootSite.createdDateTime).toLocaleString() : 'Unknown'
      };

      return {
        contents: [
          {
            uri: uri.toString(),
            text: JSON.stringify(formattedSite, null, 2),
            mimeType: "application/json"
          }
        ]
      };
    } catch (error) {
      return {
        contents: [
          {
            uri: uri.toString(),
            text: `Error accessing root SharePoint site: ${error instanceof Error ? error.message : String(error)}`,
            mimeType: "text/plain"
          }
        ]
      };
    }
  }
);

// Dynamic resource: Site structure (lists and drives)
server.resource(
  "sharepoint://sites/{siteId}/structure", 
  "Structure of a specific SharePoint site showing all document libraries and lists. Replace {siteId} with actual site ID from sharepoint://sites/all. This shows what content containers are available in the site.",
  async (uri) => {
    try {
      const uriStr = uri.toString();
      const siteIdMatch = uriStr.match(/sharepoint:\/\/sites\/([^/]+)\/structure/);
      
      if (!siteIdMatch) {
        throw new Error("Invalid site structure URI format");
      }
      
      const siteId = siteIdMatch[1];
      
      // Get both lists and drives for the site
      const [lists, drives] = await Promise.all([
        graphClient.listSiteLists(siteId),
        graphClient.listSiteDrives(siteId)
      ]);
      
      const formattedStructure = {
        siteId,
        lists: lists.map(list => ({
          id: list.id,
          name: list.displayName || list.name || 'Unnamed List',
          description: list.description || '',
          webUrl: list.webUrl || '',
          template: list.list?.template || 'Unknown',
          hidden: list.list?.hidden || false,
          type: 'list'
        })),
        drives: drives.map(drive => ({
          id: drive.id,
          name: drive.name || 'Unnamed Drive',
          description: drive.description || '',
          webUrl: drive.webUrl || '',
          driveType: drive.driveType || 'Unknown',
          owner: drive.owner?.user?.displayName || 'Unknown',
          type: 'drive'
        }))
      };

      return {
        contents: [
          {
            uri: uri.toString(),
            text: JSON.stringify(formattedStructure, null, 2),
            mimeType: "application/json"
          }
        ]
      };
    } catch (error) {
      return {
        contents: [
          {
            uri: uri.toString(),
            text: `Error accessing site structure: ${error instanceof Error ? error.message : String(error)}`,
            mimeType: "text/plain"
          }
        ]
      };
    }
  }
);

// Dynamic resource: Folder/drive contents  
server.resource(
  "sharepoint://sites/{siteId}/files/{path}",
  "Browse files and folders in a SharePoint site. Replace {siteId} with site ID and {path} with folder path (use 'root' for top level). Example: sharepoint://sites/mysite.sharepoint.com,abc123/files/Documents/Projects",
  async (uri) => {
    try {
      const uriStr = uri.toString();
      const pathMatch = uriStr.match(/sharepoint:\/\/sites\/([^/]+)\/files\/(.+)/);
      
      if (!pathMatch) {
        throw new Error("Invalid files URI format");
      }
      
      const siteId = pathMatch[1];
      let path = pathMatch[2];
      
      // Handle special case for root
      if (path === 'root') {
        path = '';
      } else {
        // Decode URI component
        path = decodeURIComponent(path);
      }
      
      const items = await graphClient.listDriveItems({
        siteId,
        path: path || undefined
      });
      
      const formattedItems = {
        siteId,
        path: path || '/',
        items: items.map(item => ({
          id: item.id,
          name: item.name,
          webUrl: item.webUrl || '',
          size: item.size || 0,
          isFolder: !!item.folder,
          isFile: !!item.file,
          mimeType: item.file?.mimeType || null,
          childCount: item.folder?.childCount || null,
          lastModifiedDateTime: item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleString() : 'Unknown',
          lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Unknown'
        }))
      };

      return {
        contents: [
          {
            uri: uri.toString(),
            text: JSON.stringify(formattedItems, null, 2),
            mimeType: "application/json"
          }
        ]
      };
    } catch (error) {
      return {
        contents: [
          {
            uri: uri.toString(),
            text: `Error accessing folder contents: ${error instanceof Error ? error.message : String(error)}`,
            mimeType: "text/plain"
          }
        ]
      };
    }
  }
);

// Dynamic resource: Search results
server.resource(
  "sharepoint://search/{query}",
  "Search results across all SharePoint content. Replace {query} with URL-encoded search terms. Example: sharepoint://search/budget%202024 or sharepoint://search/project%20proposal. Returns top 20 matching items.",
  async (uri) => {
    try {
      const uriStr = uri.toString();
      const queryMatch = uriStr.match(/sharepoint:\/\/search\/(.+)/);
      
      if (!queryMatch) {
        throw new Error("Invalid search URI format");
      }
      
      const query = decodeURIComponent(queryMatch[1]);
      
      const results = await graphClient.searchSharePoint({
        query,
        top: 20
      });
      
      const formattedResults = {
        query,
        resultCount: results.length,
        results: results.map(result => ({
          id: result.id,
          name: result.name || result.title || 'Unnamed',
          webUrl: result.webUrl || '',
          summary: result.summary || '',
          resourceType: result.resource?.["@odata.type"] || 'Unknown'
        }))
      };

      return {
        contents: [
          {
            uri: uri.toString(),
            text: JSON.stringify(formattedResults, null, 2),
            mimeType: "application/json"
          }
        ]
      };
    } catch (error) {
      return {
        contents: [
          {
            uri: uri.toString(),
            text: `Error performing search: ${error instanceof Error ? error.message : String(error)}`,
            mimeType: "text/plain"
          }
        ]
      };
    }
  }
);

server.prompt(
  "sharepoint-site-exploration-prompt",
  "A prompt to help explore SharePoint sites and their content.",
  { param: z.string().optional().describe("Not used") },
  async () => {
    return {
      messages: []
    };
  }
);

server.connect(new StdioServerTransport());