export type TUserListField = 'Email' | 'Id' | 'LoginName' | 'Title';


export interface IUserList {
  Email: string;
  Groups: {
    __deferred: {
      uri: string;
    };
  };
  Id: number;
  IsHiddenInUI: boolean;
  IsSiteAdmin: boolean;
  LoginName: string;
  PrincipalType: number;
  Title: string;
  UserId: {
    NameId: string;
    NameIdIssuer: string;
    __metadata: {
      type: string;
    }
  };
  __metadata: {
    id: string;
    type: string;
    uri: string;
  };
}
