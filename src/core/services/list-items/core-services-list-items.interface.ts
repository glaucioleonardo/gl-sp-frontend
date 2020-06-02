export interface IAttachmentServerData {
  __metadata: IAttachmentMetadata;
  FileName: string;
  ServerRelativeUrl: string;
}

interface IAttachmentMetadata {
  id: string;
  uri: string;
  type: string;
}

export interface IAttachmentData {
  id: number;
  new: boolean;
  remove: boolean;
  name: string;
  file?: File;
  url?: string;
  icon?: string;
}
