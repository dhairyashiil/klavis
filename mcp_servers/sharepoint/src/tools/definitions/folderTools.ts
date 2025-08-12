import { z } from 'zod';

export const listFoldersSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (optional for personal files)'),
  drive_id: z.string().optional().describe('Drive ID (optional for personal files)'),
  folder_id: z.string().default('root').optional().describe('Parent folder ID to list folders from (defaults to root)'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to use personal OneDrive files instead of site files'),
  limit: z
    .number()
    .min(1)
    .max(100)
    .default(20)
    .optional()
    .describe('Number of folders to retrieve (1-100)'),
});

export const createFolderSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (required for site folders)'),
  drive_id: z.string().optional().describe('Drive ID (required for site folders)'),
  parent_folder_id: z.string().default('root').optional().describe('Parent folder ID where to create the new folder'),
  folder_name: z.string().min(1).describe('Name of the folder to create'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to create folder in personal OneDrive'),
});

export const deleteFolderSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (required for site folders)'),
  drive_id: z.string().optional().describe('Drive ID (required for site folders)'),
  folder_id: z.string().describe('Folder ID to delete'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to delete from personal OneDrive'),
});

export const getFolderDetailsSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (required for site folders)'),
  drive_id: z.string().optional().describe('Drive ID (required for site folders)'),
  folder_id: z.string().describe('Folder ID to get details for'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to get details from personal OneDrive'),
  include_children: z.boolean().default(false).optional().describe('Include folder contents in the response'),
});

export const listFoldersDefinition = {
  name: 'sharepoint_list_folders',
  description:
    'List folders in a SharePoint site or personal OneDrive with optional pagination',
  inputSchema: {
    type: 'object' as const,
    properties: {
      site_id: {
        type: 'string' as const,
        description: 'SharePoint site ID (optional for personal files)',
      },
      drive_id: {
        type: 'string' as const,
        description: 'Drive ID (optional for personal files)',
      },
      folder_id: {
        type: 'string' as const,
        description: 'Parent folder ID to list folders from (defaults to root)',
        default: 'root',
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to use personal OneDrive files instead of site files',
        default: false,
      },
      limit: {
        type: 'number' as const,
        description: 'Number of folders to retrieve (1-100)',
        minimum: 1,
        maximum: 100,
        default: 20,
      },
    },
    additionalProperties: false,
  },
};

export const createFolderDefinition = {
  name: 'sharepoint_create_folder',
  description:
    'Create a new folder in SharePoint site or personal OneDrive',
  inputSchema: {
    type: 'object' as const,
    properties: {
      site_id: {
        type: 'string' as const,
        description: 'SharePoint site ID (required for site folders)',
      },
      drive_id: {
        type: 'string' as const,
        description: 'Drive ID (required for site folders)',
      },
      parent_folder_id: {
        type: 'string' as const,
        description: 'Parent folder ID where to create the new folder',
        default: 'root',
      },
      folder_name: {
        type: 'string' as const,
        description: 'Name of the folder to create',
        minLength: 1,
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to create folder in personal OneDrive',
        default: false,
      },
    },
    required: ['folder_name'],
    additionalProperties: false,
  },
};

export const deleteFolderDefinition = {
  name: 'sharepoint_delete_folder',
  description:
    'Delete an empty folder from SharePoint site or personal OneDrive',
  inputSchema: {
    type: 'object' as const,
    properties: {
      site_id: {
        type: 'string' as const,
        description: 'SharePoint site ID (required for site folders)',
      },
      drive_id: {
        type: 'string' as const,
        description: 'Drive ID (required for site folders)',
      },
      folder_id: {
        type: 'string' as const,
        description: 'Folder ID to delete',
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to delete from personal OneDrive',
        default: false,
      },
    },
    required: ['folder_id'],
    additionalProperties: false,
  },
};

export const getFolderDetailsDefinition = {
  name: 'sharepoint_get_folder_details',
  description:
    'Get detailed information about a specific folder including metadata and optional contents',
  inputSchema: {
    type: 'object' as const,
    properties: {
      site_id: {
        type: 'string' as const,
        description: 'SharePoint site ID (required for site folders)',
      },
      drive_id: {
        type: 'string' as const,
        description: 'Drive ID (required for site folders)',
      },
      folder_id: {
        type: 'string' as const,
        description: 'Folder ID to get details for',
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to get details from personal OneDrive',
        default: false,
      },
      include_children: {
        type: 'boolean' as const,
        description: 'Include folder contents in the response',
        default: false,
      },
    },
    required: ['folder_id'],
    additionalProperties: false,
  },
};
