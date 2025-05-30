import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { getAccessToken, isAuthenticated, acquireToken } from './auth.js';
import { 
  SharePointSite,
  SharePointList,
  SharePointDrive,
  SharePointDriveItem,
  SharePointSitePage,
  ListSitesQuery,
  ListDriveItemsQuery,
  GetDriveItemContentParams,
  SearchSharePointParams,
  SharePointSearchResult
} from './types.js';

const msalAuthProvider = async (done: (error: any, accessToken: string | null) => void) => {
  try {
    const token = await getAccessToken();
    done(null, token);
  } catch (error) {
    done(error, null);
  }
};

/**
 * Microsoft Graph client wrapper for SharePoint
 */
export class GraphClient {
  private client: Client;

  constructor() {
    this.client = Client.init({
      authProvider: msalAuthProvider,
    });
  }

  /**
   * Ensure the user is authenticated before calling Graph API
   */
  private async ensureAuthenticated(): Promise<void> {
    const authenticated = await isAuthenticated();
    if (!authenticated) {
      await acquireToken();
    }
  }

  // ============= SharePoint Sites Methods =============

  /**
   * List SharePoint sites that the user has access to
   */
  async listSites(query: ListSitesQuery = {}): Promise<SharePointSite[]> {
    await this.ensureAuthenticated();

    let endpoint = '/sites';
    const queryParams = new URLSearchParams();

    if (query.search) {
      queryParams.append('$search', `"${query.search}"`);
    }

    if (query.filter) {
      queryParams.append('$filter', query.filter);
    }

    if (query.top) {
      queryParams.append('$top', query.top.toString());
    }

    if (query.orderBy) {
      queryParams.append('$orderby', query.orderBy);
    }

    if (queryParams.toString()) {
      endpoint += `?${queryParams.toString()}`;
    }

    const response = await this.client.api(endpoint).get();
    return response.value;
  }

  /**
   * Get a specific SharePoint site by ID
   */
  async getSite(siteId: string): Promise<SharePointSite> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/sites/${siteId}`).get();
    return response;
  }

  /**
   * Get subsites of a SharePoint site
   */
  async getSubsites(siteId: string): Promise<SharePointSite[]> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/sites/${siteId}/sites`).get();
    return response.value;
  }

  // ============= SharePoint Lists Methods =============

  /**
   * List all lists in a SharePoint site
   */
  async listSiteLists(siteId: string): Promise<SharePointList[]> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/sites/${siteId}/lists`).get();
    return response.value;
  }

  /**
   * Get a specific list from a SharePoint site
   */
  async getSiteList(siteId: string, listId: string): Promise<SharePointList> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/sites/${siteId}/lists/${listId}`).get();
    return response;
  }

  // ============= SharePoint Drives (Document Libraries) Methods =============

  /**
   * List all drives (document libraries) in a SharePoint site
   */
  async listSiteDrives(siteId: string): Promise<SharePointDrive[]> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/sites/${siteId}/drives`).get();
    return response.value;
  }

  /**
   * Get the default drive of a SharePoint site
   */
  async getSiteDefaultDrive(siteId: string): Promise<SharePointDrive> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/sites/${siteId}/drive`).get();
    return response;
  }

  /**
   * Get a specific drive from a SharePoint site
   */
  async getSiteDrive(siteId: string, driveId: string): Promise<SharePointDrive> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/sites/${siteId}/drives/${driveId}`).get();
    return response;
  }

  // ============= SharePoint Drive Items (Files/Folders) Methods =============

  /**
   * List items in a drive or folder
   */
  async listDriveItems(query: ListDriveItemsQuery): Promise<SharePointDriveItem[]> {
    await this.ensureAuthenticated();

    let endpoint: string;
    
    if (query.driveId) {
      if (query.path) {
        endpoint = `/sites/${query.siteId}/drives/${query.driveId}/root:/${query.path}:/children`;
      } else {
        endpoint = `/sites/${query.siteId}/drives/${query.driveId}/root/children`;
      }
    } else {
      if (query.path) {
        endpoint = `/sites/${query.siteId}/drive/root:/${query.path}:/children`;
      } else {
        endpoint = `/sites/${query.siteId}/drive/root/children`;
      }
    }

    const queryParams = new URLSearchParams();

    if (query.filter) {
      queryParams.append('$filter', query.filter);
    }

    if (query.top) {
      queryParams.append('$top', query.top.toString());
    }

    if (query.orderBy) {
      queryParams.append('$orderby', query.orderBy);
    }

    if (queryParams.toString()) {
      endpoint += `?${queryParams.toString()}`;
    }

    const response = await this.client.api(endpoint).get();
    return response.value;
  }

  /**
   * Get a specific drive item by ID
   */
  async getDriveItem(siteId: string, driveId: string | undefined, itemId: string): Promise<SharePointDriveItem> {
    await this.ensureAuthenticated();

    let endpoint: string;
    if (driveId) {
      endpoint = `/sites/${siteId}/drives/${driveId}/items/${itemId}`;
    } else {
      endpoint = `/sites/${siteId}/drive/items/${itemId}`;
    }

    const response = await this.client.api(endpoint).get();
    return response;
  }

  /**
   * Get the content of a file
   */
  async getDriveItemContent(params: GetDriveItemContentParams): Promise<{ content: string; isBase64: boolean; mimeType?: string }> {
    await this.ensureAuthenticated();

    let endpoint: string;
    if (params.driveId) {
      endpoint = `/sites/${params.siteId}/drives/${params.driveId}/items/${params.itemId}/content`;
    } else {
      endpoint = `/sites/${params.siteId}/drive/items/${params.itemId}/content`;
    }

    // First get the item metadata to determine content type
    const itemEndpoint = params.driveId 
      ? `/sites/${params.siteId}/drives/${params.driveId}/items/${params.itemId}`
      : `/sites/${params.siteId}/drive/items/${params.itemId}`;
    
    const itemMetadata = await this.client.api(itemEndpoint).get();
    const mimeType = itemMetadata.file?.mimeType;

    // Get the content as a stream
    const response = await this.client.api(endpoint).getStream();
    
    // Convert stream to buffer
    const chunks: Buffer[] = [];
    for await (const chunk of response) {
      chunks.push(Buffer.from(chunk));
    }
    const buffer = Buffer.concat(chunks);

    // Determine if this should be treated as text or binary
    const isTextFile = mimeType && (
      mimeType.startsWith('text/') ||
      mimeType === 'application/json' ||
      mimeType === 'application/xml' ||
      mimeType === 'application/javascript' ||
      mimeType === 'application/typescript' ||
      mimeType.includes('xml')
    );

    if (isTextFile) {
      return {
        content: buffer.toString('utf-8'),
        isBase64: false,
        mimeType
      };
    } else {
      return {
        content: buffer.toString('base64'),
        isBase64: true,
        mimeType
      };
    }
  }

  // ============= SharePoint Search Methods =============

  /**
   * Search across SharePoint content
   */
  async searchSharePoint(params: SearchSharePointParams): Promise<SharePointSearchResult[]> {
    await this.ensureAuthenticated();

    const requestBody: any = {
      requests: [{
        entityTypes: params.entityTypes || ["listItem", "driveItem"],
        query: {
          queryString: params.query
        },
        from: 0,
        size: params.top || 25
      }]
    };

    if (params.siteId) {
      requestBody.requests[0].sharePointOneDriveOptions = {
        includeContent: "privateContent,sharedContent"
      };
    }

    const response = await this.client.api('/search/query').post(requestBody);
    
    if (response.value && response.value.length > 0 && response.value[0].hitsContainers) {
      const hits = response.value[0].hitsContainers[0]?.hits || [];
      return hits.map((hit: any) => ({
        id: hit.hitId,
        webUrl: hit.resource?.webUrl,
        name: hit.resource?.name,
        title: hit.resource?.title,
        summary: hit.summary,
        hitHighlightedSummary: hit.hitHighlightedSummary,
        resource: hit.resource
      }));
    }

    return [];
  }

  /**
   * Get root SharePoint site (organization's main site)
   */
  async getRootSite(): Promise<SharePointSite> {
    await this.ensureAuthenticated();

    const response = await this.client.api('/sites/root').get();
    return response;
  }

  // ============= SharePoint Site Pages Methods =============

  /**
   * List all site pages in a SharePoint site with canvasLayout included
   */
  async listSitePages(siteId: string, top?: number): Promise<SharePointSitePage[]> {
    await this.ensureAuthenticated();

    let endpoint = `/sites/${siteId}/pages`;
    const queryParams = new URLSearchParams();
    
    // Always expand canvasLayout to get page content
    queryParams.append('$expand', 'canvasLayout');
    
    if (top) {
      queryParams.append('$top', top.toString());
    }

    endpoint += `?${queryParams.toString()}`;

    const response = await this.client.api(endpoint).get();
    return response.value;
  }

  /**
   * Get a specific site page by ID with canvasLayout included
   */
  async getSitePage(siteId: string, pageId: string): Promise<SharePointSitePage> {
    await this.ensureAuthenticated();

    // Always expand canvasLayout to get page content
    const endpoint = `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage?$expand=canvasLayout`;

    const response = await this.client.api(endpoint).get();
    return response;
  }
}

export const graphClient = new GraphClient();