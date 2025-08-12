import { CallToolRequest } from '@modelcontextprotocol/sdk/types.js';
import { getSharePointClient, safeLog } from '../../client/sharepointClient.js';
import {
  listFoldersSchema,
  createFolderSchema,
  deleteFolderSchema,
  getFolderDetailsSchema,
} from '../definitions/folderTools.js';

export async function handleListFolders(request: CallToolRequest) {
  try {
    const args = listFoldersSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    const folderId = args.folder_id || 'root';

    let result;
    if (args.use_personal_files) {
      result = await client.getMyFiles(folderId);
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site folders');
      }
      result = await client.getFolderContents(args.site_id, args.drive_id, folderId);
    }

    
    let folders = (result.value || []).filter((item: any) => item.folder);

    
    if (args.limit) {
      folders = folders.slice(0, args.limit);
    }

    safeLog('info', `Retrieved ${folders.length} folders from SharePoint`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                folders: folders.map((folder: any) => ({
                  id: folder.id,
                  name: folder.name,
                  created: folder.createdDateTime,
                  modified: folder.lastModifiedDateTime,
                  webUrl: folder.webUrl,
                  childCount: folder.folder?.childCount || 0,
                  parentReference: folder.parentReference,
                })),
                total_count: folders.length,
                pagination: {
                  limit: args.limit || 20,
                  applied_filters: {
                    folder_id: folderId,
                    use_personal_files: args.use_personal_files || false,
                  },
                },
                api_response: 'Successfully retrieved folders from SharePoint',
              },
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleListFolders: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_list_folders',
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

export async function handleCreateFolder(request: CallToolRequest) {
  try {
    const args = createFolderSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    const parentFolderId = args.parent_folder_id || 'root';

    let result;
    if (args.use_personal_files) {
      result = await client.createMyFolder(parentFolderId, args.folder_name);
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site folders');
      }
      result = await client.createFolder(args.site_id, args.drive_id, parentFolderId, args.folder_name);
    }

    safeLog('info', `Successfully created folder: ${args.folder_name}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                folder: {
                  id: result.id,
                  name: result.name,
                  created: result.createdDateTime,
                  webUrl: result.webUrl,
                  parentReference: result.parentReference,
                },
                creation_info: {
                  folder_name: args.folder_name,
                  parent_folder_id: parentFolderId,
                  use_personal_files: args.use_personal_files || false,
                },
              },
              api_response: 'Successfully created folder in SharePoint',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleCreateFolder: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_create_folder',
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

export async function handleDeleteFolder(request: CallToolRequest) {
  try {
    const args = deleteFolderSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    if (args.use_personal_files) {
      await client.deleteMyItem(args.folder_id);
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site folders');
      }
      await client.deleteItem(args.site_id, args.drive_id, args.folder_id);
    }

    safeLog('info', `Successfully deleted folder with ID: ${args.folder_id}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                folder_id: args.folder_id,
                deleted: true,
                use_personal_files: args.use_personal_files || false,
              },
              api_response: 'Successfully deleted folder from SharePoint',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleDeleteFolder: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_delete_folder',
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

export async function handleGetFolderDetails(request: CallToolRequest) {
  try {
    const args = getFolderDetailsSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    let folderDetails;
    let folderContents = null;

    if (args.use_personal_files) {
      folderDetails = await client.get(`/me/drive/items/${args.folder_id}`);
      if (args.include_children) {
        folderContents = await client.getMyFiles(args.folder_id);
      }
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site folders');
      }
      folderDetails = await client.get(`/sites/${args.site_id}/drives/${args.drive_id}/items/${args.folder_id}`);
      if (args.include_children) {
        folderContents = await client.getFolderContents(args.site_id, args.drive_id, args.folder_id);
      }
    }

    safeLog('info', `Retrieved folder details for ID: ${args.folder_id}`);

    const responseData: any = {
      folder: {
        id: folderDetails.id,
        name: folderDetails.name,
        created: folderDetails.createdDateTime,
        modified: folderDetails.lastModifiedDateTime,
        webUrl: folderDetails.webUrl,
        size: folderDetails.size,
        childCount: folderDetails.folder?.childCount || 0,
        parentReference: folderDetails.parentReference,
        createdBy: folderDetails.createdBy,
        lastModifiedBy: folderDetails.lastModifiedBy,
      },
      request_info: {
        folder_id: args.folder_id,
        include_children: args.include_children || false,
        use_personal_files: args.use_personal_files || false,
      },
    };

    if (args.include_children && folderContents) {
      responseData.children = {
        folders: (folderContents.value || [])
          .filter((item: any) => item.folder)
          .map((folder: any) => ({
            id: folder.id,
            name: folder.name,
            modified: folder.lastModifiedDateTime,
            childCount: folder.folder?.childCount || 0,
          })),
        files: (folderContents.value || [])
          .filter((item: any) => item.file)
          .map((file: any) => ({
            id: file.id,
            name: file.name,
            size: file.size,
            modified: file.lastModifiedDateTime,
            mimeType: file.file?.mimeType,
          })),
      };
    }

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: responseData,
              api_response: 'Successfully retrieved folder details from SharePoint',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleGetFolderDetails: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_get_folder_details',
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
