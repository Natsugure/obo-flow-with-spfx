import { AadHttpClient } from "@microsoft/sp-http";

export interface IOboLoginProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  apiUrl: string;
  apiClient: AadHttpClient;
}
