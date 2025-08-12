import { z } from 'zod';

export const listProjectsSchema = z.object({
  limit: z
    .number()
    .min(1)
    .max(100)
    .default(10)
    .optional()
    .describe('Number of projects to retrieve (1-100)'),
  offset: z
    .number()
    .min(0)
    .default(0)
    .optional()
    .describe('Number of projects to skip for pagination'),
  status: z
    .enum(['active', 'completed', 'on_hold', 'cancelled'])
    .optional()
    .describe('Filter projects by status'),
  owner_id: z.string().optional().describe('Filter projects by owner user ID'),
});

export const getProjectDetailsSchema = z.object({
  project_id: z.string().describe('Project ID to retrieve details for'),
  include_tasks: z
    .boolean()
    .default(false)
    .optional()
    .describe('Include project tasks in the response'),
  include_documents: z
    .boolean()
    .default(false)
    .optional()
    .describe('Include project documents in the response'),
});

export const listProjectDocumentsSchema = z.object({
  project_id: z.string().describe('Project ID to list documents for'),
  limit: z
    .number()
    .min(1)
    .max(100)
    .default(20)
    .optional()
    .describe('Number of documents to retrieve (1-100)'),
  document_type: z
    .enum(['all', 'plans', 'reports', 'resources', 'templates'])
    .default('all')
    .optional()
    .describe('Filter documents by type'),
});

export const createProjectSchema = z.object({
  name: z.string().min(1).describe('Project name'),
  description: z.string().optional().describe('Project description'),
  owner_id: z.string().optional().describe('Project owner user ID'),
  start_date: z.string().optional().describe('Project start date (YYYY-MM-DD format)'),
  end_date: z.string().optional().describe('Project end date (YYYY-MM-DD format)'),
  status: z
    .enum(['active', 'completed', 'on_hold', 'cancelled'])
    .default('active')
    .optional()
    .describe('Project status'),
});

export const updateProjectSchema = z.object({
  project_id: z.string().describe('Project ID to update'),
  name: z.string().optional().describe('Updated project name'),
  description: z.string().optional().describe('Updated project description'),
  status: z
    .enum(['active', 'completed', 'on_hold', 'cancelled'])
    .optional()
    .describe('Updated project status'),
  start_date: z.string().optional().describe('Updated start date (YYYY-MM-DD format)'),
  end_date: z.string().optional().describe('Updated end date (YYYY-MM-DD format)'),
});

export const listProjectsDefinition = {
  name: 'sharepoint_list_projects',
  description:
    'Retrieve projects from SharePoint Project Online with optional filters for status and owner',
  inputSchema: {
    type: 'object' as const,
    properties: {
      limit: {
        type: 'number' as const,
        description: 'Number of projects to retrieve (1-100)',
        minimum: 1,
        maximum: 100,
        default: 10,
      },
      offset: {
        type: 'number' as const,
        description: 'Number of projects to skip for pagination',
        minimum: 0,
        default: 0,
      },
      status: {
        type: 'string' as const,
        description: 'Filter projects by status',
        enum: ['active', 'completed', 'on_hold', 'cancelled'],
      },
      owner_id: {
        type: 'string' as const,
        description: 'Filter projects by owner user ID',
      },
    },
    additionalProperties: false,
  },
};

export const getProjectDetailsDefinition = {
  name: 'sharepoint_get_project_details',
  description:
    'Get detailed information about a specific project including tasks and documents',
  inputSchema: {
    type: 'object' as const,
    properties: {
      project_id: {
        type: 'string' as const,
        description: 'Project ID to retrieve details for',
      },
      include_tasks: {
        type: 'boolean' as const,
        description: 'Include project tasks in the response',
        default: false,
      },
      include_documents: {
        type: 'boolean' as const,
        description: 'Include project documents in the response',
        default: false,
      },
    },
    required: ['project_id'],
    additionalProperties: false,
  },
};

export const listProjectDocumentsDefinition = {
  name: 'sharepoint_list_project_documents',
  description:
    'List documents associated with a specific project',
  inputSchema: {
    type: 'object' as const,
    properties: {
      project_id: {
        type: 'string' as const,
        description: 'Project ID to list documents for',
      },
      limit: {
        type: 'number' as const,
        description: 'Number of documents to retrieve (1-100)',
        minimum: 1,
        maximum: 100,
        default: 20,
      },
      document_type: {
        type: 'string' as const,
        description: 'Filter documents by type',
        enum: ['all', 'plans', 'reports', 'resources', 'templates'],
        default: 'all',
      },
    },
    required: ['project_id'],
    additionalProperties: false,
  },
};

export const createProjectDefinition = {
  name: 'sharepoint_create_project',
  description:
    'Create a new project in SharePoint Project Online',
  inputSchema: {
    type: 'object' as const,
    properties: {
      name: {
        type: 'string' as const,
        description: 'Project name',
        minLength: 1,
      },
      description: {
        type: 'string' as const,
        description: 'Project description',
      },
      owner_id: {
        type: 'string' as const,
        description: 'Project owner user ID',
      },
      start_date: {
        type: 'string' as const,
        description: 'Project start date (YYYY-MM-DD format)',
        pattern: '^\\d{4}-\\d{2}-\\d{2}$',
      },
      end_date: {
        type: 'string' as const,
        description: 'Project end date (YYYY-MM-DD format)',
        pattern: '^\\d{4}-\\d{2}-\\d{2}$',
      },
      status: {
        type: 'string' as const,
        description: 'Project status',
        enum: ['active', 'completed', 'on_hold', 'cancelled'],
        default: 'active',
      },
    },
    required: ['name'],
    additionalProperties: false,
  },
};

export const updateProjectDefinition = {
  name: 'sharepoint_update_project',
  description:
    'Update an existing project in SharePoint Project Online',
  inputSchema: {
    type: 'object' as const,
    properties: {
      project_id: {
        type: 'string' as const,
        description: 'Project ID to update',
      },
      name: {
        type: 'string' as const,
        description: 'Updated project name',
      },
      description: {
        type: 'string' as const,
        description: 'Updated project description',
      },
      status: {
        type: 'string' as const,
        description: 'Updated project status',
        enum: ['active', 'completed', 'on_hold', 'cancelled'],
      },
      start_date: {
        type: 'string' as const,
        description: 'Updated start date (YYYY-MM-DD format)',
        pattern: '^\\d{4}-\\d{2}-\\d{2}$',
      },
      end_date: {
        type: 'string' as const,
        description: 'Updated end date (YYYY-MM-DD format)',
        pattern: '^\\d{4}-\\d{2}-\\d{2}$',
      },
    },
    required: ['project_id'],
    additionalProperties: false,
  },
};
