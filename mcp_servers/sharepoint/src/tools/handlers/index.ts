
export * from './documentHandler.js';
export * from './folderHandler.js';
export * from './projectHandler.js';


import {
  handleListDocuments,
  handleGetDocumentContent,
  handleUploadDocument,
  handleDeleteDocument,
  handleUpdateDocument,
} from './documentHandler.js';

import {
  handleListFolders,
  handleCreateFolder,
  handleDeleteFolder,
  handleGetFolderDetails,
} from './folderHandler.js';

import {
  handleListProjects,
  handleGetProjectDetails,
  handleListProjectDocuments,
  handleCreateProject,
  handleUpdateProject,
} from './projectHandler.js';


export const documentHandlers = {
  sharepoint_list_documents: handleListDocuments,
  sharepoint_get_document_content: handleGetDocumentContent,
  sharepoint_upload_document: handleUploadDocument,
  sharepoint_delete_document: handleDeleteDocument,
  sharepoint_update_document: handleUpdateDocument,
};


export const folderHandlers = {
  sharepoint_list_folders: handleListFolders,
  sharepoint_create_folder: handleCreateFolder,
  sharepoint_delete_folder: handleDeleteFolder,
  sharepoint_get_folder_details: handleGetFolderDetails,
};


export const projectHandlers = {
  sharepoint_list_projects: handleListProjects,
  sharepoint_get_project_details: handleGetProjectDetails,
  sharepoint_list_project_documents: handleListProjectDocuments,
  sharepoint_create_project: handleCreateProject,
  sharepoint_update_project: handleUpdateProject,
};


export const allHandlers = {
  ...documentHandlers,
  ...folderHandlers,
  ...projectHandlers,
};


export const HANDLER_CATEGORIES = {
  DOCUMENTS: 'documents',
  FOLDERS: 'folders',
  PROJECTS: 'projects',
} as const;


export function getHandlersByCategory(category: keyof typeof HANDLER_CATEGORIES) {
  switch (category) {
    case 'DOCUMENTS':
      return documentHandlers;
    case 'FOLDERS':
      return folderHandlers;
    case 'PROJECTS':
      return projectHandlers;
    default:
      return {};
  }
}


export function getHandlerByName(toolName: string) {
  return allHandlers[toolName as keyof typeof allHandlers];
}


export type HandlerName = keyof typeof allHandlers;


export async function executeHandler(toolName: string, request: any) {
  const handler = getHandlerByName(toolName);
  if (!handler) {
    throw new Error(`Handler not found for tool: ${toolName}`);
  }
  return await handler(request);
}
