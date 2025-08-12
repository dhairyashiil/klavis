import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { AsyncLocalStorage } from 'async_hooks';
import { ConfidentialClientApplication, ClientCredentialRequest } from '@azure/msal-node';
import axios, { AxiosInstance, AxiosRequestConfig } from 'axios';
import dotenv from 'dotenv';

dotenv.config();

const MICROSOFT_GRAPH_API_URL = 'https://graph.microsoft.com/v1.0';

const asyncLocalStorage = new AsyncLocalStorage<{
  sharepointClient: SharePointClient;
}>();

let mcpServerInstance: Server | null = null;

export interface SharePointConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  scope?: string;
}

export class SharePointClient {
  private msalInstance: ConfidentialClientApplication;
  private httpClient: AxiosInstance;
  private accessToken: string | null = null;
  private tokenExpiry: Date | null = null;
  private config: SharePointConfig;
  private baseUrl: string;

  constructor(config: SharePointConfig, baseUrl: string = MICROSOFT_GRAPH_API_URL) {
    this.config = config;
    this.baseUrl = baseUrl;
    
    const clientConfig = {
      auth: {
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
    };

    this.msalInstance = new ConfidentialClientApplication(clientConfig);
    this.httpClient = axios.create({
      baseURL: this.baseUrl,
      timeout: 30000,
    });

    
    this.httpClient.interceptors.request.use(async (config) => {
      await this.ensureValidToken();
      config.headers.Authorization = `Bearer ${this.accessToken}`;
      return config;
    });

    
    this.httpClient.interceptors.response.use(
      (response) => response,
      (error) => {
        const errorMessage = error.response?.data?.error?.message || error.message;
        throw new Error(`SharePoint API error: ${errorMessage}`);
      }
    );
  }

  private async ensureValidToken(): Promise<void> {
    if (this.accessToken && this.tokenExpiry && new Date() < this.tokenExpiry) {
      return; 
    }

    await this.refreshToken();
  }

  private async refreshToken(): Promise<void> {
    try {
      const clientCredentialRequest: ClientCredentialRequest = {
        scopes: [this.config.scope || 'https://graph.microsoft.com/.default'],
      };

      const response = await this.msalInstance.acquireTokenByClientCredential(clientCredentialRequest);
      
      if (!response) {
        throw new Error('Failed to acquire access token');
      }

      this.accessToken = response.accessToken;
      this.tokenExpiry = response.expiresOn;
      
      safeLog('info', 'SharePoint access token refreshed successfully');
    } catch (error) {
      safeLog('error', `Failed to acquire token: ${error}`);
      throw new Error(`Failed to acquire token: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Execute a GET request to Microsoft Graph API
   */
  public async get<T = any>(endpoint: string, config?: AxiosRequestConfig): Promise<T> {
    try {
      const response = await this.httpClient.get(endpoint, config);
      return response.data;
    } catch (error) {
      throw new Error(
        `SharePoint GET error: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
    }
  }

  /**
   * Execute a POST request to Microsoft Graph API
   */
  public async post<T = any>(endpoint: string, data?: any, config?: AxiosRequestConfig): Promise<T> {
    try {
      const response = await this.httpClient.post(endpoint, data, config);
      return response.data;
    } catch (error) {
      throw new Error(
        `SharePoint POST error: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
    }
  }

  /**
   * Execute a PUT request to Microsoft Graph API
   */
  public async put<T = any>(endpoint: string, data?: any, config?: AxiosRequestConfig): Promise<T> {
    try {
      const response = await this.httpClient.put(endpoint, data, config);
      return response.data;
    } catch (error) {
      throw new Error(
        `SharePoint PUT error: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
    }
  }

  /**
   * Execute a DELETE request to Microsoft Graph API
   */
  public async delete(endpoint: string, config?: AxiosRequestConfig): Promise<void> {
    try {
      await this.httpClient.delete(endpoint, config);
    } catch (error) {
      throw new Error(
        `SharePoint DELETE error: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
    }
  }

  /**
   * Get all accessible SharePoint sites
   */
  public async getSites(): Promise<any> {
    return this.get('/sites');
  }

  /**
   * Get a specific SharePoint site by ID
   */
  public async getSiteById(siteId: string): Promise<any> {
    return this.get(`/sites/${siteId}`);
  }

  /**
   * Get drives (document libraries) for a site
   */
  public async getDrives(siteId: string): Promise<any> {
    return this.get(`/sites/${siteId}/drives`);
  }

  /**
   * Get folder contents (files and folders)
   */
  public async getFolderContents(siteId: string, driveId: string, folderId: string = 'root'): Promise<any> {
    return this.get(`/sites/${siteId}/drives/${driveId}/items/${folderId}/children`);
  }

  /**
   * Get file content
   */
  public async getFileContent(siteId: string, driveId: string, fileId: string): Promise<any> {
    return this.get(`/sites/${siteId}/drives/${driveId}/items/${fileId}/content`, {
      responseType: 'arraybuffer'
    });
  }

  /**
   * Upload a file
   */
  public async uploadFile(siteId: string, driveId: string, folderId: string, fileName: string, content: Buffer): Promise<any> {
    return this.put(
      `/sites/${siteId}/drives/${driveId}/items/${folderId}:/${fileName}:/content`,
      content,
      {
        headers: {
          'Content-Type': 'application/octet-stream',
        },
      }
    );
  }

  /**
   * Create a folder
   */
  public async createFolder(siteId: string, driveId: string, parentFolderId: string, folderName: string): Promise<any> {
    return this.post(`/sites/${siteId}/drives/${driveId}/items/${parentFolderId}/children`, {
      name: folderName,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename',
    });
  }

  /**
   * Delete an item (file or folder)
   */
  public async deleteItem(siteId: string, driveId: string, itemId: string): Promise<void> {
    return this.delete(`/sites/${siteId}/drives/${driveId}/items/${itemId}`);
  }

  /**
   * Get user's personal OneDrive
   */
  public async getMyDrive(): Promise<any> {
    return this.get('/me/drive');
  }

  /**
   * Get user's personal files
   */
  public async getMyFiles(folderId: string = 'root'): Promise<any> {
    return this.get(`/me/drive/items/${folderId}/children`);
  }

  /**
   * Upload file to personal OneDrive
   */
  public async uploadToMyDrive(folderId: string, fileName: string, content: Buffer): Promise<any> {
    return this.put(
      `/me/drive/items/${folderId}:/${fileName}:/content`,
      content,
      {
        headers: {
          'Content-Type': 'application/octet-stream',
        },
      }
    );
  }

  /**
   * Create folder in personal OneDrive
   */
  public async createMyFolder(parentFolderId: string, folderName: string): Promise<any> {
    return this.post(`/me/drive/items/${parentFolderId}/children`, {
      name: folderName,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename',
    });
  }

  /**
   * Delete item from personal OneDrive
   */
  public async deleteMyItem(itemId: string): Promise<void> {
    return this.delete(`/me/drive/items/${itemId}`);
  }

  /**
   * Test the API connection
   */
  public async testConnection(): Promise<boolean> {
    try {
      await this.get('/me');
      safeLog('info', 'SharePoint API connection test successful');
      return true;
    } catch (error) {
      safeLog('error', `Failed to connect to SharePoint API: ${error}`);
      return false;
    }
  }

  /**
   * Get current access token (for debugging)
   */
  public getAccessToken(): string | null {
    return this.accessToken;
  }

  /**
   * Check if token is valid
   */
  public isTokenValid(): boolean {
    return !!(this.accessToken && this.tokenExpiry && new Date() < this.tokenExpiry);
  }
}

function getSharePointClient(): SharePointClient {
  const store = asyncLocalStorage.getStore();
  if (!store || !store.sharepointClient) {
    throw new Error('Store not found in AsyncLocalStorage');
  }
  if (!store.sharepointClient) {
    throw new Error('SharePoint client not found in AsyncLocalStorage');
  }
  return store.sharepointClient;
}

function safeLog(
  level: 'error' | 'debug' | 'info' | 'notice' | 'warning' | 'critical' | 'alert' | 'emergency',
  data: any,
): void {
  try {
    const logData = typeof data === 'object' ? JSON.stringify(data, null, 2) : data;
    console.log(`[${level.toUpperCase()}] ${logData}`);
  } catch (error) {
    console.log(`[${level.toUpperCase()}] [LOG_ERROR] Could not serialize log data`);
  }
}

export { getSharePointClient, safeLog, asyncLocalStorage, mcpServerInstance };
