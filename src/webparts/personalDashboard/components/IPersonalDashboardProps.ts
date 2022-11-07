import { WebPartContext } from "@microsoft/sp-webpart-base";
import { GraphFI } from "@pnp/graph";

export interface IPersonalDashboardProps {
  description: string;
  graph: GraphFI;
  context: WebPartContext;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
