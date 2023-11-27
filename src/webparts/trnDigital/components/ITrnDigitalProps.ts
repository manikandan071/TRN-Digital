import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ITrnDigitalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  libraryName: string;
}
