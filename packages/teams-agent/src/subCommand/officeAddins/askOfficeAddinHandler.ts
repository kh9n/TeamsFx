import {
  AgentRequest,
} from '../../chat/agent';

import {
  SlashCommandHandlerResult,
} from '../../chat/slashCommands';

import {
  getResponseAsStringCopilotInteraction
} from "../../chat/copilotInteractions";
import { Planner } from './planner';

export async function askOfficeAddinHandler(
  request: AgentRequest
): Promise<SlashCommandHandlerResult> {

  if (!await shouldBeHandledByAskOfficeAddinHandler(request)) {
    return undefined;
  }

  let result = await Planner.getInstance().processRequest(request);

  return result ?? { chatAgentResult: { metadata: { slashCommand: 'askOfficeAddin' } }, followUp: [] };
}

async function shouldBeHandledByAskOfficeAddinHandler(request: AgentRequest): Promise<boolean> {
  const defaultSystemPrompt = `Let us start over and forget my previous input for you, as well as your answer to me. You are an expert in Office JavaScript Add-ins. The Office add-ins platform is a rich framework for building add-ins for Office applications that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Outlook, Excel, Word, PowerPoint, OneNote, and Project. Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser. Check the intention of the user and see if the ask is about Office JavaScript Add-ins. Summarize the result into "Yes" or "No", then put that with your confident score (as xx%), response strictly following this format: [Your confidence score] : [Your summary]. No need to add any explanations for your answer.`;

  const startTime = Date.now();
  // Perform the desired operation

  const copilotResponse = await getResponseAsStringCopilotInteraction(
    defaultSystemPrompt,
    request
  );

  const endTime = Date.now();
  const timeDifferenceInSeconds = Math.floor((endTime - startTime) / 1000);
  console.log(`[shouldBeHandledByAskOfficeAddinHandler] Time taken to get response from Copilot: ${timeDifferenceInSeconds} seconds. The response is: ${copilotResponse}`);

  let regex = /\d+/g;
  let match = copilotResponse.match(regex);
  let confidenceScore = match ? Number(match[0]) : 0;

  if (confidenceScore >= 80 && copilotResponse.trim().toLowerCase().indexOf("yes") >= 0) {
    return true;
  }

  return false;
}
