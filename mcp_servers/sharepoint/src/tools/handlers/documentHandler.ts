import { CallToolRequest } from '@modelcontextprotocol/sdk/types.js';
import { getSharePointClient, safeLog } from '../../client/sharepointClient.js';
import {
  listDocumentsSchema,
  getDocumentContentSchema,
  uploadDocumentSchema,
  deleteDocumentSchema,
  updateDocumentSchema,
} from '../definitions/documentTools.js';

export async function handleListDocuments(request: CallToolRequest) {
  try {
    const args = listDocumentsSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    let result;
    if (args.use_personal_files) {
      result = await client.getMyFiles(args.folder_id);
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site documents');
      }
      result = await client.getFolderContents(args.site_id, args.drive_id, args.folder_id);
    }

    
    let documents = (result.value || []).filter((item: any) => item.file);

    
    if (args.file_type && args.file_type !== 'all') {
      documents = documents.filter((doc: any) => {
        const extension = doc.name.split('.').pop()?.toLowerCase();
        switch (args.file_type) {
          case 'word':
            return ['doc', 'docx'].includes(extension || '');
          case 'excel':
            return ['xls', 'xlsx'].includes(extension || '');
          case 'powerpoint':
            return ['ppt', 'pptx'].includes(extension || '');
          case 'pdf':
            return extension === 'pdf';
          case 'text':
            return ['txt', 'md', 'csv'].includes(extension || '');
          case 'image':
            return ['jpg', 'jpeg', 'png', 'gif', 'bmp'].includes(extension || '');
          default:
            return true;
        }
      });
    }

    
    if (args.limit) {
      documents = documents.slice(0, args.limit);
    }

    safeLog('info', `Retrieved ${documents.length} documents from SharePoint`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                documents: documents.map((doc: any) => ({
                  id: doc.id,
                  name: doc.name,
                  size: doc.size,
                  created: doc.createdDateTime,
                  modified: doc.lastModifiedDateTime,
                  webUrl: doc.webUrl,
                  downloadUrl: doc['@microsoft.graph.downloadUrl'],
                  mimeType: doc.file?.mimeType,
                })),
                total_count: documents.length,
                pagination: {
                  limit: args.limit || 20,
                  applied_filters: {
                    file_type: args.file_type || 'all',
                    folder_id: args.folder_id || 'root',
                    use_personal_files: args.use_personal_files || false,
                  },
                },
                api_response: 'Successfully retrieved documents from SharePoint',
              },
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleListDocuments: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_list_documents',
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

export async function handleGetDocumentContent(request: CallToolRequest) {
  try {
    const args = getDocumentContentSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    let content;
    if (args.use_personal_files) {
      content = await client.get(`/me/drive/items/${args.file_id}/content`, {
        responseType: args.return_format === 'base64' ? 'arraybuffer' : 'text',
      });
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site documents');
      }
      content = await client.getFileContent(args.site_id, args.drive_id, args.file_id);
    }

    let formattedContent;
    switch (args.return_format) {
      case 'base64':
        formattedContent = Buffer.isBuffer(content) 
          ? content.toString('base64')
          : Buffer.from(content).toString('base64');
        break;
      case 'url':
        
        const fileInfo = args.use_personal_files
          ? await client.get(`/me/drive/items/${args.file_id}`)
          : await client.get(`/sites/${args.site_id}/drives/${args.drive_id}/items/${args.file_id}`);
        formattedContent = fileInfo['@microsoft.graph.downloadUrl'];
        break;
      default:
        formattedContent = typeof content === 'string' ? content : content.toString();
    }

    safeLog('info', `Retrieved document content for file ID: ${args.file_id}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                file_id: args.file_id,
                content: formattedContent,
                format: args.return_format || 'text',
                content_length: typeof formattedContent === 'string' 
                  ? formattedContent.length 
                  : 'Binary content',
              },
              api_response: 'Successfully retrieved document content from SharePoint',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleGetDocumentContent: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_get_document_content',
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

export async function handleUploadDocument(request: CallToolRequest) {
  try {
    const args = uploadDocumentSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    const fileBuffer = Buffer.from(args.content, 'base64');
    const folderId = args.folder_id || 'root';

    let result;
    if (args.use_personal_files) {
      result = await client.uploadToMyDrive(folderId, args.file_name, fileBuffer);
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site uploads');
      }
      result = await client.uploadFile(args.site_id, args.drive_id, folderId, args.file_name, fileBuffer);
    }

    safeLog('info', `Successfully uploaded document: ${args.file_name}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                file: {
                  id: result.id,
                  name: result.name,
                  size: result.size,
                  created: result.createdDateTime,
                  webUrl: result.webUrl,
                  downloadUrl: result['@microsoft.graph.downloadUrl'],
                },
                upload_info: {
                  file_name: args.file_name,
                  folder_id: folderId,
                  overwrite: args.overwrite || false,
                  use_personal_files: args.use_personal_files || false,
                },
              },
              api_response: 'Successfully uploaded document to SharePoint',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleUploadDocument: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_upload_document',
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


export async function handleDeleteDocument(request: CallToolRequest) {
  try {
    const args = deleteDocumentSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    if (args.use_personal_files) {
      await client.deleteMyItem(args.file_id);
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site documents');
      }
      await client.deleteItem(args.site_id, args.drive_id, args.file_id);
    }

    safeLog('info', `Successfully deleted document with ID: ${args.file_id}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                file_id: args.file_id,
                deleted: true,
                use_personal_files: args.use_personal_files || false,
              },
              api_response: 'Successfully deleted document from SharePoint',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleDeleteDocument: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_delete_document',
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

export async function handleUpdateDocument(request: CallToolRequest) {
  try {
    const args = updateDocumentSchema.parse(request.params.arguments || {});
    const client = getSharePointClient();

    const fileBuffer = Buffer.from(args.content, 'base64');

    let result;
    if (args.use_personal_files) {
      result = await client.put(`/me/drive/items/${args.file_id}/content`, fileBuffer, {
        headers: {
          'Content-Type': 'application/octet-stream',
        },
      });
    } else {
      if (!args.site_id || !args.drive_id) {
        throw new Error('site_id and drive_id are required for SharePoint site documents');
      }
      result = await client.put(`/sites/${args.site_id}/drives/${args.drive_id}/items/${args.file_id}/content`, fileBuffer, {
        headers: {
          'Content-Type': 'application/octet-stream',
        },
      });
    }

    safeLog('info', `Successfully updated document with ID: ${args.file_id}`);

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: true,
              data: {
                file: {
                  id: result.id,
                  name: result.name,
                  size: result.size,
                  modified: result.lastModifiedDateTime,
                  webUrl: result.webUrl,
                },
                update_info: {
                  file_id: args.file_id,
                  use_personal_files: args.use_personal_files || false,
                },
              },
              api_response: 'Successfully updated document in SharePoint',
            },
            null,
            2,
          ),
        },
      ],
    };
  } catch (error) {
    safeLog('error', `Error in handleUpdateDocument: ${error}`);
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(
            {
              success: false,
              error: error instanceof Error ? error.message : 'Unknown error occurred',
              tool: 'sharepoint_update_document',
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
