import { CancellationToken, ChatContext, ChatFollowup } from 'vscode';
import { AgentRequest, ITeamsChatAgentResult } from '../../../chat/agent';
import {
  getResponseAsStringCopilotInteraction,
  verbatimCopilotInteraction
} from "../../../chat/copilotInteractions";
import { SlashCommandHandlerResult } from '../../../chat/slashCommands';
import { SampleProvider } from '../samples/sampleProvider';
import { ISkill } from './iSkill'; // Add the missing import statement


export class CodeGenerator implements ISkill {
  name: string;
  promptForAdditionalInput: string;
  capability: string;

  constructor() {
    this.name = "Code Generator";
    this.capability = "How to automate a process?";
    this.promptForAdditionalInput = "If this is the case, briefly descript what the process user should do, to automate using Office JavaScript add-in or api. Set the description as part of the additional input. Think how to break down that task into smaller step by step guidances, given explainations for each step, as rest part of additional input. ";
  }

  public canInvoke(capability: string, additionalInput: string): boolean {
    if (this.capability.indexOf(capability) > -1) {
      return true;
    }

    return false;
  }

  followUp?: ((result: ITeamsChatAgentResult, context: ChatContext, token: CancellationToken) => ChatFollowup[] | undefined) | undefined;

  public async invoke(additionalInput: string, request: AgentRequest): Promise<SlashCommandHandlerResult> {
    let preScanningPrompt = `The task is about: ${request.userPrompt}, and this is the detail of task: ${additionalInput}

    Break down the task into step by step sub tasks could performed by Office add-in JavaScript APIs, and list them below. Reference to the API reference:
    Add-ins API reference for Excel: https://learn.microsoft.com/en-us/javascript/api/excel?view=excel-js-preview&viewFallbackFrom=word-js-preview)
    Add-ins API reference for Word: https://learn.microsoft.com/en-us/javascript/api/word?view=word-js-preview&viewFallbackFrom=outlook-js-preview)
    Add-ins API reference for Outlook: https://learn.microsoft.com/en-us/javascript/api/outlook?view=outlook-js-preview&viewFallbackFrom=common-js-preview)
    Add-ins API reference for PowerPoint: https://learn.microsoft.com/en-us/javascript/api/outlook?view=outlook-js-preview&viewFallbackFrom=powerpoint-js-preview)
    Add-ins API reference for OneNote: https://learn.microsoft.com/en-us/javascript/api/onenote?view=onenote-js-1.1)
    Add-ins API reference for Visio: https://learn.microsoft.com/en-us/javascript/api/visio?view=visio-js-1.1&viewFallbackFrom=common-js-preview)
    Add-ins API reference for Common: https://learn.microsoft.com/en-us/javascript/api/office?view=common-js-preview)

    And give a full list of Office Add-in API Class, properties of Class, and methods of Class that you would use to perform each sub task. The output of following format: [Your confidence score] [line break] Sub-tasks-start: [Your sub task description] Sub-tasks-end [line break] Objects/Properties/Method-start: [line break] [Your list of Class, properties, and methods] [line break] Objects/Properties/Method-end. No need to add any explanations for your answer.

    Alternatively, if the user's ask is not clear, and you can't make a recommendation based on the context to cover those missed information, you should stop processing and ask for clarification. Indicate if you can't make a recommendation based on the context by return following context as output, formatted as:
    Ask-for-clarification-start: [line break] I can't make a recommendation based on the context to cover following missed information. [The list of missed information]. [line break] Ask-for-clarification-end.

    This is a sample output for the ask to clarify:
    Ask-for-clarification-start:
    I can't make a recommendation based on the context to cover following missed information.
    - The stock price API endpoint.
    - The prediction algorithm or model.
    Ask-for-clarification-end.

    This is a sample output for the ask no need ask for clarification:
    95%
    Sub-tasks-start:
    1. Use Office JavaScript API to create a new Excel worksheet.
    2. Use a stock price API to fetch the last two week's MSFT stock price.
    3. Import the fetched data into the Excel worksheet.
    4. Use a prediction algorithm or model to predict the next trading day's price.
    5. Display the predicted price in the Excel worksheet.
    Sub-tasks-end
    Objects/Properties/Method-start:
    1. Create a new Excel worksheet:
      - Class: 'Excel.Workbook'
      - Property: 'worksheets'
      - Method: 'add'

    2. Fetch the last two week's MSFT stock price:
      - Global Function: 'fetch'

    3. Import the fetched data into the Excel worksheet:
      - Class: 'Excel.Worksheet'
      - Method: 'getRange'
      - Class: 'Excel.Range'
      - Property: 'values'

    4. Predict the next trading day's price:
      - Global Function: 'Array.prototype.reduce'
      - Property: 'Array.prototype.length'

    5. Display the predicted price in the Excel worksheet:
      - Class: 'Excel.Worksheet'
      - Method: 'getRange'
      - Class: 'Excel.Range'
      - Property: 'values'
    Objects/Properties/Method-end

    Think that step by step.
    `;

    let startTime = Date.now();
    // Perform the desired operation

    let copilotResponse = await getResponseAsStringCopilotInteraction(
      preScanningPrompt,
      request
    );

    let endTime = Date.now();
    let timeDifferenceInSeconds = Math.floor((endTime - startTime) / 1000);
    console.log(`[generateCode - Task breakdown/snipper] Time taken to get response from Copilot: ${timeDifferenceInSeconds} seconds.  The response is: ${copilotResponse}`);

    let beginIndex = copilotResponse.indexOf("Ask-for-clarification-start:") + "Ask-for-clarification-start:".length;
    let endIndex = copilotResponse.indexOf("Ask-for-clarification-end");
    if (beginIndex > -1 && endIndex > beginIndex) {
      // build the hints for the user to provide the missing information
      const missingInfo = copilotResponse.substring(beginIndex, endIndex);
      const missingInfoHints = `
The task is about: ${request.userPrompt}<br>
It could be detailed write down as: ${additionalInput}<br>
However in order to continue to process, I need **more detail or clarification**:<br>
${missingInfo}<br>
Please provide the missing information in chat window, and then I can continue to process.<br>
      `;

      // The control will give back to user, and we've done our part..
      request.response.markdown(missingInfoHints);

      return {
        chatAgentResult: { metadata: { slashCommand: "" } },
        followUp: [],
      };
    }

    beginIndex = copilotResponse.indexOf("Sub-tasks-start:") + "Sub-tasks-start:".length;
    endIndex = copilotResponse.indexOf("Sub-tasks-end");
    let subTasks = beginIndex > -1 && endIndex > beginIndex ? copilotResponse.substring(beginIndex, endIndex) : additionalInput;
    let firstRoundGeneratePrompt = `
    The user ask is: ${request.userPrompt}. And that could be break down into a few steps:\r\n${subTasks}\r\n
    Generate Office JavaScript add-in function or code snippet according to the steps you listed above. You should strickly following those rules when you generate code:
    1. Use real existing or product code.
    2. Do not use placeholder.
    3. Do not use pseudo code or hypothetical code.
    4. Do not use hypothetical API
    5. Do not use hypothetical service endpoint.
    6. Instead, use well-known service, algorithm, or solutions as recommendation to cover uncleared details. Generate code using that recommendation.
    7. Use real world data.
    8. OK to use placeholder for credentials, API keys, tokens or other sensitive information.
    9. OK to reference simliar code and transpile into javascript from the existing code base.
    10. Implement all the steps you listed above, do not leave EMPTY implementation.
    `;

    // let us generate a list of Class-Property pairs, and Class-Method pairs in the format of JSON object, from the objectsPropertiesMethods string
    beginIndex = copilotResponse.indexOf("Objects\/Properties\/Method-start:") + "Objects\/Properties\/Method-start:".length;
    endIndex = copilotResponse.indexOf("Objects\/Properties\/Method-end");
    if (beginIndex > -1 && endIndex > beginIndex) {
      let objectsPropertiesMethods = copilotResponse.substring(beginIndex, endIndex);
      let classMethodPairs: { class: string; method: string; }[] = [];
      let classPropertyPairs: { class: string; property: string; }[] = [];

      let matches = objectsPropertiesMethods.split("\n").filter((item: string) => item.trim() !== "");

      let classNames = "";
      matches?.forEach(match => {
        if (match.indexOf("Class") > -1) {
          classNames = match.replace(/- (Class): '/g, "").replace(/'/g, "").trim();

          // TODO: This is a hack should be find a solution to fix
          classNames = classNames.replace("Excel.run", "").replace(",", "").trim();
        } else if (match.indexOf("Property") > -1) {
          classPropertyPairs.push({ class: classNames, property: match.replace(/- (Property): '/g, "").replace(/'/g, "").trim() });
        } else if (match.indexOf("Method") > -1) {
          classMethodPairs.push({ class: classNames, method: match.replace(/- (Method): '/g, "").replace(/'/g, "").trim() });
        }
      });

      // Now we have the classMethodPairs and classPropertyPairs, we can query the code snippet examples from the SampleProvider
      let codeSnippets: string[] = [];
      classMethodPairs.forEach(pair => {
        let snippets = SampleProvider.getInstance().getAPISampleCodes(pair.class, pair.method);
        snippets.forEach((sample, api) => {
          const intro = `- The code sample of Class: ${pair.class} and Method: ${api} for the scenario: ${sample.scenario} listed below: ${sample.sample}, API reference: ${sample.docLink}`;
          codeSnippets.push(intro);
        });
      });
      classPropertyPairs.forEach(pair => {
        let snippets = SampleProvider.getInstance().getAPISampleCodes(pair.class, pair.property);
        snippets.forEach((sample, api) => {
          const intro = `- The code sample of Class: ${pair.class} and Property: ${api} for the scenario: ${sample.scenario} listed below: ${sample.sample}, API reference: ${sample.docLink}`;
          codeSnippets.push(intro);
        });
      });

      if (codeSnippets.length > 0) {
        let codeSnippet = codeSnippets.toString().replace(/,/g, "\r\n");
        firstRoundGeneratePrompt = firstRoundGeneratePrompt.concat(`\r\nHere are some code snippets example for you to reference:\n\n${codeSnippet}\n\n`);
      }
    };

    startTime = Date.now();
    // Perform the desired operation

    copilotResponse = await getResponseAsStringCopilotInteraction(
      firstRoundGeneratePrompt,
      request
    );

    endTime = Date.now();
    timeDifferenceInSeconds = Math.floor((endTime - startTime) / 1000);
    console.log(`[generateCode - first round generation] Time taken to get response from Copilot: ${timeDifferenceInSeconds} seconds.  The response is: ${copilotResponse}`);

    const selfReflectionPrompt = `
    You're a professional in Office JavaScript Add-ins developers with a lot of experience on JavaScript, CSS, HTML, popular algrithom, and Office Add-ins API. The user is a junior engineer do not have much of experience on JavaScript, CSS, HTML, popular algrithom, and Office Add-ins API. You're asked to generate code for the user's ask, please reply with clear code structure and detail explanations.

    The user ask is: ${request.userPrompt}. And that could be break down into following steps: ${subTasks}.

    The following is the code snippet you generated:
    ${copilotResponse}

    Code above have issues, fix them and re-generate the code snippet for the user's ask. The issues are listed below:
    1. Using codes or libraries only availabe in node environment. For example, code like "require", "import", or library like "express". Those should not be used in Office JavaScript Add-in.
    2. Asynchronicity issue. For example, the code is not using the "await" keyword to call the async function. Or function is not marked as async.
    3. Context.sync() is called in the right place
    4. The context.sync() is called in a loop.
    5. Upper case set to the first letter of the enumeration, variable and function name, which should be lower case.
    6. The generate code is for Office Script. Should generate for Office JavaScript Add-in.
    7. The code is not following the best practice of Office JavaScript Add-in development.
    8. Not all desired steps are covered.

    For multiple code snippets generated for different steps, if reasonable, wrap them in a one single method.

    For the output, you should strictly following the following format:
    [Your confidence score] : [Explain how to get that confident score]
    The ask is: ${request.userPrompt}
    And it could be break down into following steps: ${subTasks}
    [Your code snippet]
    [Explaination of code snippet].
    `;

    startTime = Date.now();
    // Perform the desired operation

    let verbResponse = await verbatimCopilotInteraction(
      selfReflectionPrompt,
      request
    );

    endTime = Date.now();
    timeDifferenceInSeconds = Math.floor((endTime - startTime) / 1000);
    console.log(`[generateCode - self-refelection] Time taken to get response from Copilot: ${timeDifferenceInSeconds} seconds.`);

    return {
      chatAgentResult: { metadata: { slashCommand: "" } },
      followUp: [],
    };
  }
}
