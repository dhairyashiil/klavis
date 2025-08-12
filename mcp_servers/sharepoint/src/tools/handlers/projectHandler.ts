import { CallToolRequest } from '@modelcontextprotocol/sdk/types.js';
import { getSharePointClient, safeLog } from '../../client/sharepointClient.js';
import {
  listProjectsSchema,
  getProjectDetailsSchema,
  listProjectDocumentsSchema,
  createProjectSchema,
  updateProjectSchema,
} from '../definitions/projectTools.js';

export async function handleListProjects(request: CallToolRequest) {
  try {
    const args = listProjectsSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    
    
    let endpoint = '/sites/root/projects'; 
    const queryParams = new URLSearchParams();

    if (args.limit) queryParams.append('$top', args.limit.toString());
    if (args.offset) queryParams.append('$skip', args.offset.toString());
    if (args.status) queryParams.append('$filter', `status eq '${args.status}'`);
    if (args.owner_id) queryParams.append('$filter', `ownerId eq '${args.owner_id}'`);

    if (queryParams.toString()) {
      endpoint += `?${queryParams.toString()}`;
    }

    const result = await client.get(endpoint);

    
    let projects = result.value || [];
    if (args.status) {
      projects = projects.filter((project: any) => project.status === args.status);
    }
    if (args.owner_id) {
      projects = projects.filter((project: any) => project.ownerId === args.owner_id);
    }

    safeLog('info', `Retrieved ${projects.length} projects from SharePoint Project Online`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                projects: projects.map((project: any) => ({
                  id: project.id,
                  name: project.name,
                  description: project.description,
                  status: project.status,
                  owner_id: project.ownerId,
                  owner_name: project.ownerName,
                  start_date: project.startDate,
                  end_date: project.endDate,
                  created: project.createdDate,
                  modified: project.modifiedDate,
                  progress: project.percentComplete || 0,
                  priority: project.priority,
                })),
                total_count: projects.length,
                pagination: {
                  limit: args.limit || 10,
                  offset: args.offset || 0,
                },
                applied_filters: {
                  status: args.status,
                  owner_id: args.owner_id,
                },
                api_response: 'Successfully retrieved projects from SharePoint Project Online',
              },
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleListProjects: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_list_projects',
            },
            null,
            2,
          ),
        },
      ],
      isError: true,
    };
  }
}

export async function handleGetProjectDetails(request: CallToolRequest) {
  try {
    const args = getProjectDetailsSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    
    const projectDetails = await client.get(`/sites/root/projects/${args.project_id}`);

    const responseData: any = {
      project: {
        id: projectDetails.id,
        name: projectDetails.name,
        description: projectDetails.description,
        status: projectDetails.status,
        owner_id: projectDetails.ownerId,
        owner_name: projectDetails.ownerName,
        start_date: projectDetails.startDate,
        end_date: projectDetails.endDate,
        created: projectDetails.createdDate,
        modified: projectDetails.modifiedDate,
        progress: projectDetails.percentComplete || 0,
        priority: projectDetails.priority,
        budget: projectDetails.budget,
        team_members: projectDetails.teamMembers || [],
      },
      request_info: {
        project_id: args.project_id,
        include_tasks: args.include_tasks || false,
        include_documents: args.include_documents || false,
      },
    };

    
    if (args.include_tasks) {
      try {
        const tasksResult = await client.get(`/sites/root/projects/${args.project_id}/tasks`);
        responseData.tasks = (tasksResult.value || []).map((task: any) => ({
          id: task.id,
          name: task.name,
          status: task.status,
          assigned_to: task.assignedTo,
          start_date: task.startDate,
          due_date: task.dueDate,
          progress: task.percentComplete || 0,
          priority: task.priority,
        }));
      } catch (taskError) {
        safeLog('warning', `Could not retrieve tasks for project ${args.project_id}: ${taskError}`);
        responseData.tasks = [];
      }
    }

    
    if (args.include_documents) {
      try {
        const documentsResult = await client.get(`/sites/root/projects/${args.project_id}/documents`);
        responseData.documents = (documentsResult.value || []).map((doc: any) => ({
          id: doc.id,
          name: doc.name,
          type: doc.documentType,
          size: doc.size,
          created: doc.createdDateTime,
          modified: doc.lastModifiedDateTime,
          webUrl: doc.webUrl,
        }));
      } catch (docError) {
        safeLog('warning', `Could not retrieve documents for project ${args.project_id}: ${docError}`);
        responseData.documents = [];
      }
    }

    safeLog('info', `Retrieved project details for ID: ${args.project_id}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: responseData,
              api_response: 'Successfully retrieved project details from SharePoint Project Online',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleGetProjectDetails: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_get_project_details',
            },
            null,
            2,
          ),
        },
      ],
      isError: true,
    };
  }
}

export async function handleListProjectDocuments(request: CallToolRequest) {
  try {
    const args = listProjectDocumentsSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    let endpoint = `/sites/root/projects/${args.project_id}/documents`;
    const queryParams = new URLSearchParams();

    if (args.limit) queryParams.append('$top', args.limit.toString());
    if (args.document_type && args.document_type !== 'all') {
      queryParams.append('$filter', `documentType eq '${args.document_type}'`);
    }

    if (queryParams.toString()) {
      endpoint += `?${queryParams.toString()}`;
    }

    const result = await client.get(endpoint);

    let documents = result.value || [];

    
    if (args.document_type && args.document_type !== 'all') {
      documents = documents.filter((doc: any) => doc.documentType === args.document_type);
    }

    safeLog('info', `Retrieved ${documents.length} documents for project ${args.project_id}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                project_id: args.project_id,
                documents: documents.map((doc: any) => ({
                  id: doc.id,
                  name: doc.name,
                  type: doc.documentType,
                  size: doc.size,
                  created: doc.createdDateTime,
                  modified: doc.lastModifiedDateTime,
                  webUrl: doc.webUrl,
                  downloadUrl: doc.downloadUrl,
                  version: doc.version,
                  checkoutStatus: doc.checkoutStatus,
                })),
                total_count: documents.length,
                pagination: {
                  limit: args.limit || 20,
                  applied_filters: {
                    document_type: args.document_type || 'all',
                  },
                },
                api_response: 'Successfully retrieved project documents from SharePoint',
              },
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleListProjectDocuments: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_list_project_documents',
            },
            null,
            2,
          ),
        },
      ],
      isError: true,
    };
  }
}

export async function handleCreateProject(request: CallToolRequest) {
  try {
    const args = createProjectSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    const projectData = {
      name: args.name,
      description: args.description || '',
      ownerId: args.owner_id,
      startDate: args.start_date,
      endDate: args.end_date,
      status: args.status || 'active',
      createdDate: new Date().toISOString(),
    };

    const result = await client.post('/sites/root/projects', projectData);

    safeLog('info', `Successfully created project: ${args.name}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                project: {
                  id: result.id,
                  name: result.name,
                  description: result.description,
                  status: result.status,
                  owner_id: result.ownerId,
                  start_date: result.startDate,
                  end_date: result.endDate,
                  created: result.createdDate,
                },
                creation_info: {
                  name: args.name,
                  description: args.description,
                  owner_id: args.owner_id,
                  status: args.status || 'active',
                },
              },
              api_response: 'Successfully created project in SharePoint Project Online',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleCreateProject: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_create_project',
            },
            null,
            2,
          ),
        },
      ],
      isError: true,
    };
  }
}

export async function handleUpdateProject(request: CallToolRequest) {
  try {
    const args = updateProjectSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    
    const updateData: any = {
      modifiedDate: new Date().toISOString(),
    };

    if (args.name) updateData.name = args.name;
    if (args.description !== undefined) updateData.description = args.description;
    if (args.status) updateData.status = args.status;
    if (args.start_date) updateData.startDate = args.start_date;
    if (args.end_date) updateData.endDate = args.end_date;

    const result = await client.put(`/sites/root/projects/${args.project_id}`, updateData);

    safeLog('info', `Successfully updated project with ID: ${args.project_id}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                project: {
                  id: result.id,
                  name: result.name,
                  description: result.description,
                  status: result.status,
                  start_date: result.startDate,
                  end_date: result.endDate,
                  modified: result.modifiedDate,
                },
                update_info: {
                  project_id: args.project_id,
                  updated_fields: Object.keys(updateData).filter(key => key !== 'modifiedDate'),
                },
              },
              api_response: 'Successfully updated project in SharePoint Project Online',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleUpdateProject: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_update_project',
            },
            null,
            2,
          ),
        },
      ],
      isError: true,
    };
  }
}
