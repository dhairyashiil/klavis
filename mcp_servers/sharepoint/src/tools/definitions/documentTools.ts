import { z } from 'zod';

export const listDocumentsSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (optional for personal files)'),
  drive_id: z.string().optional().describe('Drive ID (optional for personal files)'),
  folder_id: z.string().default('root').optional().describe('Folder ID to list documents from (defaults to root)'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to use personal OneDrive files instead of site files'),
  limit: z
    .number()
    .min(1)
    .max(100)
    .default(20)
    .optional()
    .describe('Number of documents to retrieve (1-100)'),
  file_type: z
    .enum(['all', 'word', 'excel', 'powerpoint', 'pdf', 'text', 'image'])
    .default('all')
    .optional()
    .describe('Filter documents by file type'),
});

export const getDocumentContentSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (required for site files)'),
  drive_id: z.string().optional().describe('Drive ID (required for site files)'),
  file_id: z.string().describe('File ID to retrieve content from'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to use personal OneDrive files'),
  return_format: z
    .enum(['text', 'base64', 'url'])
    .default('text')
    .optional()
    .describe('Format to return content in'),
});

export const uploadDocumentSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (required for site files)'),
  drive_id: z.string().optional().describe('Drive ID (required for site files)'),
  folder_id: z.string().default('root').optional().describe('Folder ID where to upload the document'),
  file_name: z.string().min(1).describe('Name of the file to upload'),
  content: z.string().describe('File content (base64 encoded for binary files)'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to upload to personal OneDrive'),
  overwrite: z.boolean().default(false).optional().describe('Whether to overwrite existing file'),
});

export const deleteDocumentSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (required for site files)'),
  drive_id: z.string().optional().describe('Drive ID (required for site files)'),
  file_id: z.string().describe('File ID to delete'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to delete from personal OneDrive'),
});

export const updateDocumentSchema = z.object({
  site_id: z.string().optional().describe('SharePoint site ID (required for site files)'),
  drive_id: z.string().optional().describe('Drive ID (required for site files)'),
  file_id: z.string().describe('File ID to update'),
  content: z.string().describe('Updated file content (base64 encoded for binary files)'),
  use_personal_files: z.boolean().default(false).optional().describe('Whether to update in personal OneDrive'),
});

export const listDocumentsDefinition = {
  name: 'sharepoint_list_documents',
  description:
    'List documents in a SharePoint folder or document library with optional filtering and pagination',
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
        description: 'Folder ID to list documents from (defaults to root)',
        default: 'root',
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to use personal OneDrive files instead of site files',
        default: false,
      },
      limit: {
        type: 'number' as const,
        description: 'Number of documents to retrieve (1-100)',
        minimum: 1,
        maximum: 100,
        default: 20,
      },
      file_type: {
        type: 'string' as const,
        description: 'Filter documents by file type',
        enum: ['all', 'word', 'excel', 'powerpoint', 'pdf', 'text', 'image'],
        default: 'all',
      },
    },
    additionalProperties: false,
  },
};

export const getDocumentContentDefinition = {
  name: 'sharepoint_get_document_content',
  description:
    'Retrieve the content of a specific document from SharePoint or OneDrive',
  inputSchema: {
    type: 'object' as const,
    properties: {
      site_id: {
        type: 'string' as const,
        description: 'SharePoint site ID (required for site files)',
      },
      drive_id: {
        type: 'string' as const,
        description: 'Drive ID (required for site files)',
      },
      file_id: {
        type: 'string' as const,
        description: 'File ID to retrieve content from',
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to use personal OneDrive files',
        default: false,
      },
      return_format: {
        type: 'string' as const,
        description: 'Format to return content in',
        enum: ['text', 'base64', 'url'],
        default: 'text',
      },
    },
    required: ['file_id'],
    additionalProperties: false,
  },
};

export const uploadDocumentDefinition = {
  name: 'sharepoint_upload_document',
  description:
    'Upload a new document to SharePoint site or personal OneDrive',
  inputSchema: {
    type: 'object' as const,
    properties: {
      site_id: {
        type: 'string' as const,
        description: 'SharePoint site ID (required for site files)',
      },
      drive_id: {
        type: 'string' as const,
        description: 'Drive ID (required for site files)',
      },
      folder_id: {
        type: 'string' as const,
        description: 'Folder ID where to upload the document',
        default: 'root',
      },
      file_name: {
        type: 'string' as const,
        description: 'Name of the file to upload',
        minLength: 1,
      },
      content: {
        type: 'string' as const,
        description: 'File content (base64 encoded for binary files)',
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to upload to personal OneDrive',
        default: false,
      },
      overwrite: {
        type: 'boolean' as const,
        description: 'Whether to overwrite existing file',
        default: false,
      },
    },
    required: ['file_name', 'content'],
    additionalProperties: false,
  },
};

export const deleteDocumentDefinition = {
  name: 'sharepoint_delete_document',
  description:
    'Delete a document from SharePoint site or personal OneDrive',
  inputSchema: {
    type: 'object' as const,
    properties: {
      site_id: {
        type: 'string' as const,
        description: 'SharePoint site ID (required for site files)',
      },
      drive_id: {
        type: 'string' as const,
        description: 'Drive ID (required for site files)',
      },
      file_id: {
        type: 'string' as const,
        description: 'File ID to delete',
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to delete from personal OneDrive',
        default: false,
      },
    },
    required: ['file_id'],
    additionalProperties: false,
  },
};

export const updateDocumentDefinition = {
  name: 'sharepoint_update_document',
  description:
    'Update the content of an existing document in SharePoint or OneDrive',
  inputSchema: {
    type: 'object' as const,
    properties: {
      site_id: {
        type: 'string' as const,
        description: 'SharePoint site ID (required for site files)',
      },
      drive_id: {
        type: 'string' as const,
        description: 'Drive ID (required for site files)',
      },
      file_id: {
        type: 'string' as const,
        description: 'File ID to update',
      },
      content: {
        type: 'string' as const,
        description: 'Updated file content (base64 encoded for binary files)',
      },
      use_personal_files: {
        type: 'boolean' as const,
        description: 'Whether to update in personal OneDrive',
        default: false,
      },
    },
    required: ['file_id', 'content'],
    additionalProperties: false,
  },
};
