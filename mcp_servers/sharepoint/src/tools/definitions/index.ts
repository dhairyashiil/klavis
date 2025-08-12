
export * from './documentTools.js';
export * from './folderTools.js';
export * from './projectTools.js';


import {
  listDocumentsDefinition,
  getDocumentContentDefinition,
  uploadDocumentDefinition,
  deleteDocumentDefinition,
  updateDocumentDefinition,
} from './documentTools.js';

import {
  listFoldersDefinition,
  createFolderDefinition,
  deleteFolderDefinition,
  getFolderDetailsDefinition,
} from './folderTools.js';

import {
  listProjectsDefinition,
  getProjectDetailsDefinition,
  listProjectDocumentsDefinition,
  createProjectDefinition,
  updateProjectDefinition,
} from './projectTools.js';


export const documentToolDefinitions = [
  listDocumentsDefinition,
  getDocumentContentDefinition,
  uploadDocumentDefinition,
  deleteDocumentDefinition,
  updateDocumentDefinition,
];


export const folderToolDefinitions = [
  listFoldersDefinition,
  createFolderDefinition,
  deleteFolderDefinition,
  getFolderDetailsDefinition,
];


export const projectToolDefinitions = [
  listProjectsDefinition,
  getProjectDetailsDefinition,
  listProjectDocumentsDefinition,
  createProjectDefinition,
  updateProjectDefinition,
];


export const allToolDefinitions = [
  ...documentToolDefinitions,
  ...folderToolDefinitions,
  ...projectToolDefinitions,
];


export const TOOL_NAMES = {
  
  LIST_DOCUMENTS: 'sharepoint_list_documents',
  GET_DOCUMENT_CONTENT: 'sharepoint_get_document_content',
  UPLOAD_DOCUMENT: 'sharepoint_upload_document',
  DELETE_DOCUMENT: 'sharepoint_delete_document',
  UPDATE_DOCUMENT: 'sharepoint_update_document',
  
  
  LIST_FOLDERS: 'sharepoint_list_folders',
  CREATE_FOLDER: 'sharepoint_create_folder',
  DELETE_FOLDER: 'sharepoint_delete_folder',
  GET_FOLDER_DETAILS: 'sharepoint_get_folder_details',
  
  
  LIST_PROJECTS: 'sharepoint_list_projects',
  GET_PROJECT_DETAILS: 'sharepoint_get_project_details',
  LIST_PROJECT_DOCUMENTS: 'sharepoint_list_project_documents',
  CREATE_PROJECT: 'sharepoint_create_project',
  UPDATE_PROJECT: 'sharepoint_update_project',
} as const;


export const TOOL_CATEGORIES = {
  DOCUMENTS: 'documents',
  FOLDERS: 'folders',
  PROJECTS: 'projects',
} as const;


export function getToolsByCategory(category: keyof typeof TOOL_CATEGORIES) {
  switch (category) {
    case 'DOCUMENTS':
      return documentToolDefinitions;
    case 'FOLDERS':
      return folderToolDefinitions;
    case 'PROJECTS':
      return projectToolDefinitions;
    default:
      return [];
  }
}


export function getToolByName(toolName: string) {
  return allToolDefinitions.find(tool => tool.name === toolName);
}
