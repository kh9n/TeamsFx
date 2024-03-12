import { AgentRequest } from "../../chat/agent";
import { SlashCommand, SlashCommandHandlerResult } from "../../chat/slashCommands";
import { askOfficeAddinHandler } from "./askOfficeAddinHandler";

const officeAddinCommandName = "askOfficeAddin";
export const OFFICEADDIN_COMMAND_ID = 'teamsAgent.askOfficeAddin';

export function getOfficeAddinCommand(): SlashCommand {
  return [officeAddinCommandName,
    {
      shortDescription: `Describe what we can do for Office Add-in development`,
      longDescription: `Describe what we can do for Office Add-in development, currently we support giving suggestions on scaffolding a template, giving code template, and what next you can do.`,
      intentDescription: '',
      handler: (request: AgentRequest) => addAskHandler(request)
    }];
}

async function addAskHandler(request: AgentRequest): Promise<SlashCommandHandlerResult> {
  return await askOfficeAddinHandler(request);
}


export async function officeAddinCommand(folderOrSample: string) {
  return;
}
