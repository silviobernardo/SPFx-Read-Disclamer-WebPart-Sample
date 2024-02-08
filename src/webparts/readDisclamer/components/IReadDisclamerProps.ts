import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISiteUserInfo } from '@pnp/sp/site-users/types';

export interface IReadDisclamerWebPartProps {
  documentTitle: string;
  storageList: string;
  acknowledgementLabel: string;
  acknowledgementMessage: string;
  readMessage: string;
  themeVariant: IReadonlyTheme | undefined;
  configured: boolean;
  context: WebPartContext;
}

export interface IReadDisclamerProps extends IReadDisclamerWebPartProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  currentUser: ISiteUserInfo | undefined;
  userDisplayName: string;
} 
