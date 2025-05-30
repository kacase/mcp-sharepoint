import { z } from "zod";

// Schema for SharePoint Site
export const SharePointSiteSchema = z.object({
  id: z.string(),
  name: z.string().optional(),
  displayName: z.string().optional(),
  description: z.string().optional(),
  webUrl: z.string(),
  createdDateTime: z.string().optional(),
  lastModifiedDateTime: z.string().optional(),
  siteCollection: z.object({
    hostname: z.string().optional()
  }).optional(),
  root: z.object({}).optional()
});

export type SharePointSite = z.infer<typeof SharePointSiteSchema>;

// Schema for SharePoint List
export const SharePointListSchema = z.object({
  id: z.string(),
  name: z.string().optional(),
  displayName: z.string().optional(),
  description: z.string().optional(),
  webUrl: z.string().optional(),
  createdDateTime: z.string().optional(),
  lastModifiedDateTime: z.string().optional(),
  list: z.object({
    template: z.string().optional(),
    hidden: z.boolean().optional()
  }).optional()
});

export type SharePointList = z.infer<typeof SharePointListSchema>;

// Schema for SharePoint Drive (Document Library)
export const SharePointDriveSchema = z.object({
  id: z.string(),
  name: z.string().optional(),
  description: z.string().optional(),
  webUrl: z.string().optional(),
  createdDateTime: z.string().optional(),
  lastModifiedDateTime: z.string().optional(),
  driveType: z.string().optional(),
  owner: z.object({
    user: z.object({
      displayName: z.string().optional()
    }).optional()
  }).optional()
});

export type SharePointDrive = z.infer<typeof SharePointDriveSchema>;

// Schema for SharePoint Drive Item (File/Folder)
export const SharePointDriveItemSchema = z.object({
  id: z.string(),
  name: z.string(),
  webUrl: z.string().optional(),
  createdDateTime: z.string().optional(),
  lastModifiedDateTime: z.string().optional(),
  size: z.number().optional(),
  folder: z.object({
    childCount: z.number().optional()
  }).optional(),
  file: z.object({
    mimeType: z.string().optional(),
    hashes: z.object({
      sha1Hash: z.string().optional()
    }).optional()
  }).optional(),
  parentReference: z.object({
    driveId: z.string().optional(),
    path: z.string().optional()
  }).optional(),
  createdBy: z.object({
    user: z.object({
      displayName: z.string().optional()
    }).optional()
  }).optional(),
  lastModifiedBy: z.object({
    user: z.object({
      displayName: z.string().optional()
    }).optional()
  }).optional()
});

export type SharePointDriveItem = z.infer<typeof SharePointDriveItemSchema>;

// Schema for listing sites query parameters
export const ListSitesQuerySchema = z.object({
  search: z.string().optional().describe("Search term to filter sites"),
  top: z.number().int().positive().optional().describe("Maximum number of sites to return"),
  filter: z.string().optional().describe("OData filter expression"),
  orderBy: z.string().optional().describe("Order by expression")
});

export type ListSitesQuery = z.infer<typeof ListSitesQuerySchema>;

// Schema for listing drive items query parameters
export const ListDriveItemsQuerySchema = z.object({
  siteId: z.string().describe("SharePoint site ID"),
  driveId: z.string().optional().describe("Drive ID (optional, uses default drive if not specified)"),
  path: z.string().optional().describe("Folder path (optional, defaults to root)"),
  top: z.number().int().positive().optional().describe("Maximum number of items to return"),
  filter: z.string().optional().describe("OData filter expression"),
  orderBy: z.string().optional().describe("Order by expression")
});

export type ListDriveItemsQuery = z.infer<typeof ListDriveItemsQuerySchema>;

// Schema for getting drive item content
export const GetDriveItemContentSchema = z.object({
  siteId: z.string().describe("SharePoint site ID"),
  driveId: z.string().optional().describe("Drive ID (optional, uses default drive if not specified)"),
  itemId: z.string().describe("Drive item ID")
});

export type GetDriveItemContentParams = z.infer<typeof GetDriveItemContentSchema>;

// Schema for searching SharePoint content
export const SearchSharePointSchema = z.object({
  query: z.string().describe("Search query"),
  siteId: z.string().optional().describe("Limit search to specific site"),
  top: z.number().int().positive().optional().describe("Maximum number of results to return"),
  entityTypes: z.array(z.enum(["listItem", "driveItem", "site", "list"])).optional().describe("Types of entities to search")
});

export type SearchSharePointParams = z.infer<typeof SearchSharePointSchema>;

// Schema for SharePoint Search Result
export const SharePointSearchResultSchema = z.object({
  id: z.string(),
  webUrl: z.string().optional(),
  name: z.string().optional(),
  title: z.string().optional(),
  summary: z.string().optional(),
  hitHighlightedSummary: z.string().optional(),
  resource: z.object({
    "@odata.type": z.string().optional(),
    id: z.string().optional(),
    name: z.string().optional(),
    webUrl: z.string().optional()
  }).optional()
});

export type SharePointSearchResult = z.infer<typeof SharePointSearchResultSchema>;

// Schema for SharePoint Site Page Canvas Layout
export const WebPartSchema = z.object({
  "@odata.type": z.string().optional(),
  id: z.string().optional(),
  innerHtml: z.string().optional()
});

export const ColumnSchema = z.object({
  id: z.string().optional(),
  width: z.number().optional(),
  webparts: z.array(WebPartSchema).optional()
});

export const HorizontalSectionSchema = z.object({
  layout: z.string().optional(),
  id: z.string().optional(),
  emphasis: z.string().optional(),
  columns: z.array(ColumnSchema).optional()
});

export const CanvasLayoutSchema = z.object({
  horizontalSections: z.array(HorizontalSectionSchema).optional()
});

// Schema for SharePoint Site Page
export const SharePointSitePageSchema = z.object({
  id: z.string(),
  name: z.string().optional(),
  title: z.string().optional(),
  description: z.string().optional(),
  webUrl: z.string().optional(),
  createdDateTime: z.string().optional(),
  lastModifiedDateTime: z.string().optional(),
  eTag: z.string().optional(),
  pageLayout: z.enum(["microsoftReserved", "article", "home", "unknownFutureValue"]).optional(),
  promotionKind: z.enum(["microsoftReserved", "page", "newsPost", "unknownFutureValue"]).optional(),
  showComments: z.boolean().optional(),
  showRecommendedPages: z.boolean().optional(),
  thumbnailWebUrl: z.string().optional(),
  createdBy: z.object({
    user: z.object({
      displayName: z.string().optional(),
      email: z.string().optional()
    }).optional()
  }).optional(),
  lastModifiedBy: z.object({
    user: z.object({
      displayName: z.string().optional(),
      email: z.string().optional()
    }).optional()
  }).optional(),
  publishingState: z.object({
    level: z.string().optional(),
    versionId: z.string().optional()
  }).optional(),
  contentType: z.object({
    id: z.string().optional(),
    name: z.string().optional()
  }).optional(),
  parentReference: z.object({
    listId: z.string().optional(),
    siteId: z.string().optional()
  }).optional(),
  reactions: z.object({}).optional(),
  canvasLayout: CanvasLayoutSchema.optional()
});

export type SharePointSitePage = z.infer<typeof SharePointSitePageSchema>;