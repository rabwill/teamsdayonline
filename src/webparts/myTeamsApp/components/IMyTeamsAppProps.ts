import { IMicrosoftTeams } from "@microsoft/sp-webpart-base";


export interface IMyTeamsAppProps {
  description: string;
  teamsContext?: IMicrosoftTeams;
}
