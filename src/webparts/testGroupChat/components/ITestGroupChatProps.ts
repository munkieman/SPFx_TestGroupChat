import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITestGroupChatProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  owners?: any[]; // Optional property to store owners
  context: WebPartContext;
  currentUserEmail?: string; // Optional property to store the current user
}
