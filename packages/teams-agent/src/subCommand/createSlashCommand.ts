import * as fs from "fs-extra";
import * as os from 'os';
import * as path from "path";
import * as tmp from "tmp";
import * as vscode from "vscode";
import { AgentRequest } from "../chat/agent";
import { getResponseAsStringCopilotInteraction, parseCopilotResponseMaybeWithStrJson, verbatimCopilotInteraction } from "../chat/copilotInteractions";
import { SlashCommand, SlashCommandHandlerResult } from "../chat/slashCommands";
import { ProjectMetadata, matchProject } from "../projectMatch";
import { SampleUrlInfo, fetchOnlineSampleConfig } from '../sample';
import { buildFileTree, downloadSampleFiles, getSampleFileInfo, modifyFile } from "../util";

const createCommandName = "create";
export const CREATE_SAMPLE_COMMAND_ID = 'teamsAgent.createSample';

export function getCreateCommand(): SlashCommand {
  return [createCommandName,
    {
      shortDescription: `Describe what kind of app you want to create in Teams`,
      longDescription: `Describe what kind of app you want to create in Teams`,
      intentDescription: '',
      handler: (request: AgentRequest) => createHandler(request)
    }];
}

async function createHandler(request: AgentRequest): Promise<SlashCommandHandlerResult> {
  if (['word', 'excel', 'powerpoint'].some(substring => request.userPrompt.toLowerCase().includes(substring.toLowerCase()))) {
    const matchedSamples: any[] = [];
    const objectJson = '{"Annotation":"Represents an annotation attached to a paragraph.","AnnotationCollection":"Contains a collection of Annotation objects.","Body":"Represents the body of a document or a section.","Border":"Represents the Border object for text, a paragraph, or a table.","BorderCollection":"Represents the collection of border styles.","CheckboxContentControl":"The data specific to content controls of type CheckBox.","Comment":"Represents a comment in the document.","CommentCollection":"Contains a collection of Comment objects.","CommentContentRange":"Specifies the comment\'s content range.","CommentReply":"Represents a comment reply in the document.","CommentReplyCollection":"Contains a collection of CommentReply objects. Represents all comment replies in one comment thread.","ContentControl":"Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, and checkbox content controls are supported.","ContentControlCollection":"Contains a collection of ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text and plain text content controls are supported.","CritiqueAnnotation":"Represents an annotation wrapper around critique displayed in the document.","CustomProperty":"Represents a custom property.","CustomPropertyCollection":"Contains the collection of CustomProperty objects.","CustomXmlPart":"Represents a custom XML part.","CustomXmlPartCollection":"Contains the collection of CustomXmlPart objects.","CustomXmlPartScopedCollection":"Contains the collection of CustomXmlPart objects with a specific namespace.","Document":"The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.","DocumentCreated":"The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.","DocumentProperties":"Represents document properties.","Field":"Represents a field.","FieldCollection":"Contains a collection of Field objects.","Font":"Represents a font.","InlinePicture":"Represents an inline picture.","InlinePictureCollection":"Contains a collection of InlinePicture objects.","List":"Contains a collection of Paragraph objects.","ListCollection":"Contains a collection of List objects.","ListItem":"Represents the paragraph list item format.","ListLevel":"Represents a list level.","ListLevelCollection":"Contains a collection of ListLevel objects.","ListTemplate":"Represents a ListTemplate.","NoteItem":"Represents a footnote or endnote.","NoteItemCollection":"Contains a collection of NoteItem objects.","Paragraph":"Represents a single paragraph in a selection, range, content control, or document body.","ParagraphCollection":"Contains a collection of Paragraph objects.","ParagraphFormat":"Represents a style of paragraph in a document.","Range":"Represents a contiguous area in a document.","RangeCollection":"Contains a collection of Range objects.","SearchOptions":"Specifies the options to be included in a search operation. To learn more about how to use search options in the Word JavaScript APIs, read Use search options to find text in your Word add-in.","Section":"Represents a section in a Word document.","SectionCollection":"Contains the collection of the document\'s Section objects.","Setting":"Represents a setting of the add -in.","SettingCollection":"Contains the collection of Setting objects.","Shading":"Represents the shading object.","Style":"Represents a style in a Word document.","StyleCollection":"Contains a collection of Style objects.","Table":"Represents a table in a Word document.","TableBorder":"Specifies the border style.","TableCell":"Represents a table cell in a Word document.","TableCellCollection":"Contains the collection of the document\'s TableCell objects.","TableCollection":"Contains the collection of the document\'s Table objects.","TableRow":"Represents a row in a Word document.","TableRowCollection":"Contains the collection of the document\'s TableRow objects.","TableStyle":"Represents the TableStyle object.","TrackedChange":"Represents a tracked change in a Word document.","TrackedChangeCollection":"Contains a collection of TrackedChange."}';
    const parsedObjectDescription = JSON.parse(objectJson);

    const generateProjectPrompt = `
        # Role
        I want you act as an expert in Office JavaScript add-in development area.You are also an advisor for Office add-in developers.

        # Instructions
        - Given the Office JavaScript add-in developer's request, please follow below to help determine the information about generating an JavaScript add-in project.
        - You should interpret the intention of developer's request as an ask to generate an Office JavaScript add-in project. And polish user input into some sentences if necessary.
        - You should go through the following steps silently, and only reply to user with a JSON result in each step. Do not explain why for your answer.

        - Suggest an platform for the add-in project.There are 3 options: Word, Excel, PowerPoint.If you can't determine, just say All.
        - You should base on your understanding of developer intent and the capabilities of Word, Excel, PowerPoint to make the suggestion.
        - Remember it as "PLATFORM".

        - Suggest an add-in type.You have 3 options: taskpane, content, custom function. You should notice Word doesn't have content type, and only Excel has custom function type. Remember it as "TYPE".

        - You should then base on the "PLATFORM" information and add-in developer asks to suggest one or a set of specific Office JavaScript API objects that are related.
        - You should analyze the API objects typical user cases or capabilities of their related UI features to suggest the most relevant ones.
        - The suggested API objects should not be too general such as "Document", "Workbook", "Presentation".
        - The suggested API objects should be from the list inside "API objects list".
        - The "API objects list" is a JSON object with a list of Office JavaScript API objects and their descriptions. The "API obejcts list" is as follows: ${JSON.stringify(parsedObjectDescription)}
        - You should give at most 3 relevant objects.Remember it as "APISET".

        - Provide some detailed summary about why you make the suggestions in above steps. Remember it as "SUMMARY".
        ` ;

    const addinPlatfromTypeAPIResponse = await getResponseAsStringCopilotInteraction(generateProjectPrompt, request);
    if (addinPlatfromTypeAPIResponse) {
      // request.response.markdown(`${addinPlatfromTypeAPIResponse}\n`);
      const responseJson = parseCopilotResponseMaybeWithStrJson(addinPlatfromTypeAPIResponse);
      if (responseJson) {
        // const projectTemplateZipFile = "https://github.com/OfficeDev/Office-Addin-TaskPane/archive/master.zip";
        // let response: any;
        // try {
        //   response = await fetch(projectTemplateZipFile, { method: "GET" });
        // } catch (e: any) {
        //   throw new Error("OfficeAddinGenerator " + e);
        // }
        // const ObjArrayWithApis = (responseJson.PLATFORM === 'Word') ? await import("./APIsWithDesciption_Word_v2_CC_Comment.json") : await import("./APIsWithDescription_Excel_Chart.json");
        // var objApisMap = new Map();
        // for (var idx in ObjArrayWithApis) {
        //   const objAndApiListJson = ObjArrayWithApis[idx];
        //   const keys = Object.keys(objAndApiListJson.apiList);
        //   objApisMap.set(objAndApiListJson.object.toLowerCase(), JSON.stringify(keys));
        // }
        const apiObjectsStr = Array.isArray(responseJson.APISET) ? responseJson.APISET.map((api: string) => `${api}`).join(", ") : '';
        // var objsFromResponse = Array.isArray(responseJson.APISET) ? Array.from(responseJson.APISET) : [];
        // request.response.markdown(`${responseJson.SUMMARY}\n`);

        // //var objsFromResponseSet = new Set(objsFromResponse);
        // var apiListStr = ""
        // for (var objFromResponseIdx in objsFromResponse) {
        //   var objFromResponse = objsFromResponse[objFromResponseIdx].toLowerCase();
        //   if (objApisMap.has(objFromResponse)) {
        //     const apiList = objApisMap.get(objFromResponse);
        //     apiListStr += objApisMap.get(objFromResponse) + ",";
        //     console.log("apiListStr " + apiListStr);
        //   }
        // }
        let codeMath = "";
        const generateCodePrompt = `
        # Role
        I want you act as an expert in Office JavaScript add-in development area. You are also an advisor for Office add-in developers.

        # Instructions
        - You should help generate some Office JavaScript API call examples based on user request.
        - The generated method must start with 'export async function' keyword.
        - The generated method should contain a meaningful function name and a runnable code snippet with its own context.
        - The generated method should have a try catch block to handle the exception.
        - Each generated method should contain Word.run, Excel.run or PowerPoint.run logic.
        - Each generated method should not have any passed in parameters. The necessary parameters should be defined inside the method.
        - The generated method for each object should contain loading properties, get and set properties and some method calls. All the properties, method calls should be existing on this object or related with it.
        - Remember to strictly reference the "API list" to generate the code. The "API list" is as follows: ${getApiListStringByObject(apiObjectsStr.split(', '))}.
        - If the userPrompt includes add or insert keywords, your generated code should contain insert or add method calls.
        `;

        const userRequestBackup = request.userPrompt;
        request.userPrompt = ` Please generate one method for each ${apiObjectsStr} ${responseJson.PLATFORM} JavaScript API object.`;
        while (codeMath === "") {
          const generatedCodeResponse = await getResponseAsStringCopilotInteraction(generateCodePrompt, request);
          if (generatedCodeResponse) {
            const quoteChar = '```';
            //const regex = new RegExp(`${quoteChar}(.*?)${quoteChar}`, 'g');
            const regex = /```javascript([\s\S]*?)```/g;
            const matches = [...generatedCodeResponse.matchAll(regex)];
            codeMath = matches.map((match) => match[1]).join('\n');

            console.log(codeMath);
          }
        }

        request.userPrompt = userRequestBackup;
        let codeMath2 = "";
        while (codeMath2 === "") {
          const generatedCodeResponse2 = await getResponseAsStringCopilotInteraction(generateCodePrompt, request);
          if (generatedCodeResponse2) {
            const quoteChar = '```';
            //const regex = new RegExp(`${quoteChar}(.*?)${quoteChar}`, 'g');
            const regex = /```javascript([\s\S]*?)```/g;
            const matches = [...generatedCodeResponse2.matchAll(regex)];
            codeMath2 = matches.map((match) => match[1]).join('\n');
            console.log(codeMath2);
          }
        }


        const wxpSampleURLInfo: SampleUrlInfo = {
          owner: "GavinGu07",
          repository: "Office-Add-in-Templates",
          ref: "main",
          dir: String(responseJson.PLATFORM)
        };
        const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(wxpSampleURLInfo, 2);
        const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
        const nodes = await buildFileTree(fileUrlPrefix, samplePaths, tempFolder, wxpSampleURLInfo.dir, 2, 20);
        request.response.filetree(nodes, vscode.Uri.file(path.join(tempFolder, wxpSampleURLInfo.dir)));

        const srcRoot = os.homedir();
        const defaultTargetFolder = srcRoot ? path.join(srcRoot, "Office-Add-in") : '';
        console.log("defaultTargetFolder: " + defaultTargetFolder);

        request.response.markdown(`Do you want to create your add-in project at the default location ${defaultTargetFolder}?\n`);

        const folder = path.join(tempFolder, wxpSampleURLInfo.dir);
        const codeMathNew1 = correctEnumSpelling(codeMath);
        const codeMathNew2 = correctEnumSpelling(codeMath2);

        await modifyFile(folder, codeMathNew1);
        await modifyFile(folder, codeMathNew2);

        request.response.button({
          command: CREATE_SAMPLE_COMMAND_ID,
          arguments: [folder, defaultTargetFolder],
          title: vscode.l10n.t('Create at the default location')
        });

        request.response.button({
          command: CREATE_SAMPLE_COMMAND_ID,
          arguments: [folder, ''],
          title: vscode.l10n.t('Create at a different location')
        });
        // request.response.markdown(`${codeMathc}\n`);


        //console.log(generatedCodeResponse)
        //request.response.markdown(`${generatedCodeResponse}\n`);

        // const followUpPrompt = `
        // # Role
        // I want you act as an expert in Office JavaScript add-in development area.You are also an advisor for Office add-in developers.

        // # Instructions
        // - You should provide three follow-up actions to help the user further develop the ${responseJson.PLATFORM} add-in.
        // - Each follow-up suggestion should have less than 10 words.
        // ` ;
        // const followUpResponse = await getResponseAsStringCopilotInteraction(followUpPrompt, request);
        // if (followUpResponse) {
        //   console.log(`${followUpResponse}\n`);
        // }

      }
    }
    const NextStepFix: vscode.ChatFollowup = {
      prompt: "Fix the errors in my code",
      command: "fix",
      label: vscode.l10n.t("Fix the errors in my code"),
    };
    const NextStepGenerate: vscode.ChatFollowup = {
      prompt: "Generate more code",
      command: "",
      label: vscode.l10n.t("Generate more code"),
    };
    return { chatAgentResult: { slashCommand: 'create' }, followUp: [NextStepFix, NextStepGenerate] };
  }

  const matchedResult = await matchProject(request);

  if (matchedResult.length === 0) {
    request.response.markdown(vscode.l10n.t("Sorry, I can't help with that right now.\n"));
    return { chatAgentResult: { slashCommand: '' }, followUp: [] };
  }
  if (matchedResult.length === 1) {
    const firstMatch = matchedResult[0];
    await describeProject(firstMatch, request);
    if (firstMatch.type === 'sample') {
      const folder = await showFileTree(firstMatch, request);
      request.response.button({
        command: CREATE_SAMPLE_COMMAND_ID,
        arguments: [folder],
        title: vscode.l10n.t('Scaffold this sample')
      });
    } else if (firstMatch.type === 'template') {
      request.response.button({
        command: "fx-extension.create",
        arguments: ["CopilotChat", firstMatch.data],
        title: vscode.l10n.t('Create this template')
      });
    }

    return { chatAgentResult: { slashCommand: 'create' }, followUp: [] };
  } else {
    request.response.markdown(`I found ${matchedResult.slice(0, 3).length} projects that match your description.\n`);
    for (const project of matchedResult.slice(0, 3)) {
      const introduction = await getResponseAsStringCopilotInteraction(
        `You are an advisor for Teams App developers. You need to describe the project based on name and description field of user's JSON content. You should control the output between 30 and 40 words.`,
        request
      );
      request.response.markdown(`- ${project.name}: ${introduction}\n`);
      if (project.type === 'sample') {
        request.response.button({
          command: CREATE_SAMPLE_COMMAND_ID,
          arguments: [project],
          title: vscode.l10n.t('Scaffold this sample')
        });
      } else if (project.type === 'template') {
        request.response.button({
          command: "fx-extension.create",
          arguments: ["CopilotChat", project.data],
          title: vscode.l10n.t('Create this template')
        });
      }
    }
    return { chatAgentResult: { slashCommand: 'create' }, followUp: [] };
  }
}

async function describeProject(projectMetadata: ProjectMetadata, request: AgentRequest): Promise<void> {
  const originPrompt = request.userPrompt;
  request.userPrompt = `The project you are looking for is '${JSON.stringify(projectMetadata)}'.`;
  await verbatimCopilotInteraction(
    `You are an advisor for Teams App developers. You need to describe the project based on name and description field of user's JSON content. You should control the output between 50 and 80 words.`,
    request
  );
  request.userPrompt = originPrompt;
}

async function showFileTree(projectMetadata: ProjectMetadata, request: AgentRequest): Promise<string> {
  request.response.markdown(vscode.l10n.t('\nHere is the files of the sample project.'));
  const downloadUrlInfo = await getSampleDownloadUrlInfo(projectMetadata.id);
  const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(downloadUrlInfo, 2);
  const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
  const nodes = await buildFileTree(fileUrlPrefix, samplePaths, tempFolder, downloadUrlInfo.dir, 2, 20);
  request.response.filetree(nodes, vscode.Uri.file(path.join(tempFolder, downloadUrlInfo.dir)));
  return path.join(tempFolder, downloadUrlInfo.dir);
}

async function getSampleDownloadUrlInfo(sampleId: string): Promise<SampleUrlInfo> {
  const sampleConfig = await fetchOnlineSampleConfig();
  const sample = sampleConfig.samples.find((sample) => sample.id === sampleId);
  let downloadUrlInfo = {
    owner: "OfficeDev",
    repository: "TeamsFx-Samples",
    ref: "dev",
    dir: sampleId,
  };
  if (sample && sample["downloadUrlInfo"]) {
    downloadUrlInfo = sample["downloadUrlInfo"] as SampleUrlInfo;
  }
  return downloadUrlInfo;
}

export async function createCommand(folderOrSample: string | ProjectMetadata, dstPath: string) {
  // Let user choose the project folder
  let folderChoice: string | undefined = undefined;
  if (dstPath === "") {
    if (vscode.workspace.workspaceFolders !== undefined && vscode.workspace.workspaceFolders.length > 0) {
      folderChoice = await vscode.window.showQuickPick(["Current workspace", "Browse..."]);
      if (!folderChoice) {
        return;
      }
      if (folderChoice === "Current workspace") {
        dstPath = vscode.workspace.workspaceFolders[0].uri.fsPath;
      }
    }
    if (dstPath === "") {
      const customFolder = await vscode.window.showOpenDialog({
        title: "Choose where to save your project",
        openLabel: "Select Folder",
        canSelectFiles: false,
        canSelectFolders: true,
        canSelectMany: false,
      });
      if (!customFolder) {
        return;
      }
      dstPath = customFolder[0].fsPath;
    }
  }
  try {
    if (typeof folderOrSample === "string") {
      await fs.copy(folderOrSample, dstPath);
    } else {
      const downloadUrlInfo = await getSampleDownloadUrlInfo(folderOrSample.id);
      const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(downloadUrlInfo, 2);
      await downloadSampleFiles(fileUrlPrefix, samplePaths, dstPath, downloadUrlInfo.dir, 2, 20);
    }
    if (folderChoice !== "Current workspace") {
      void vscode.commands.executeCommand(
        "vscode.openFolder",
        vscode.Uri.file(dstPath),
      );
    } else {
      vscode.window.showInformationMessage('Project is created in current workspace.');
      // vscode.commands.executeCommand('workbench.view.extension.teamsfx');
    }
  } catch (error) {
    console.error('Error copying files:', error);
    vscode.window.showErrorMessage('Project cannot be created.');
  }
}

export function getApiListStringByObject(obs: string[]): string {
  let apiList = "";
  apiWithDescriptionJson.map((item) => {
    if (isRelated(obs, item.object)) {
      apiList += "// " + item.object + ":\n";
      apiList += item.apiList;
      apiList += "\n";
    }
  });
  console.log(apiList);
  return apiList;
}

function isRelated(strings: string[], target: string): boolean {
  return strings.some(item => item.toLocaleLowerCase().includes(target.toLocaleLowerCase()) || target.toLocaleLowerCase().includes(item.toLocaleLowerCase()));
}

var apiWithDescriptionJson = [
  {
    "object": "body",
    "apiList": "body.contentControls: Gets the collection of rich text content control objects in the body.\nbody.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\nbody.endnotes: Gets the collection of endnotes in the body.\nbody.fields: Gets the collection of field objects in the body.\nbody.font: Gets the text format of the body. Use this to get and set font name, size, color and other properties.\nbody.footnotes: Gets the collection of footnotes in the body.\nbody.inlinePictures: Gets the collection of InlinePicture objects in the body. The collection doesn't include floating images.\nbody.lists: Gets the collection of list objects in the body.\nbody.paragraphs: Gets the collection of paragraph objects in the body.\nbody.parentBody: Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an ItemNotFound error if there isn't a parent body.\nbody.parentBodyOrNullObject: Gets the parent body of the body. For example, a table cell body's parent body could be a header. If there isn't a parent body, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nbody.parentContentControl: Gets the content control that contains the body. Throws an ItemNotFound error if there isn't a parent content control.\nbody.parentContentControlOrNullObject: Gets the content control that contains the body. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nbody.parentSection: Gets the parent section of the body. Throws an ItemNotFound error if there isn't a parent section.\nbody.parentSectionOrNullObject: Gets the parent section of the body. If there isn't a parent section, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nbody.style: Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the 'styleBuiltIn' property.\nbody.styleBuiltIn: Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the 'style' property.\nbody.tables: Gets the collection of table objects in the body.\nbody.text: Gets the text of the body. Use the insertText method to insert text.\nbody.type: Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Additional types \u2018Footnote\u2019, \u2018Endnote\u2019, and \u2018NoteItem\u2019 are supported in WordAPIOnline 1.1 and later.\nbody.clear(): Clears the contents of the body object. The user can perform the undo operation on the cleared content.\nbody.getComments(): Gets comments associated with the body.\nbody.getContentControls(options): Gets the currently supported content controls in the body.\nbody.getHtml(): Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use Body.getOoxml() and convert the returned XML to HTML.\nbody.getOoxml(): Gets the OOXML (Office Open XML) representation of the body object.\nbody.getRange(rangeLocation): Gets the whole body, or the starting or ending point of the body, as a range.\nbody.getReviewedText(changeTrackingVersion): Gets reviewed text based on ChangeTrackingVersion selection.\nbody.getReviewedText(changeTrackingVersionString): Gets reviewed text based on ChangeTrackingVersion selection.\nbody.getTrackedChanges(): Gets the collection of the TrackedChange objects in the body.\nbody.insertBreak(breakType, insertLocation): Inserts a break at the specified location in the main document.\nbody.insertContentControl(contentControlType): Wraps the Body object with a content control.\nbody.insertFileFromBase64(base64File, insertLocation): Inserts a document into the body at the specified location.\nbody.insertHtml(html, insertLocation): Inserts HTML at the specified location.\nbody.insertInlinePictureFromBase64(base64EncodedImage, insertLocation): Inserts a picture into the body at the specified location.\nbody.insertOoxml(ooxml, insertLocation): Inserts OOXML at the specified location.\nbody.insertParagraph(paragraphText, insertLocation): Inserts a paragraph at the specified location.\nbody.insertTable(rowCount, columnCount, insertLocation, values): Inserts a table with the specified number of rows and columns.\nbody.insertText(text, insertLocation): Inserts text into the body at the specified location.\nbody.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nbody.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nbody.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nbody.search(searchText, searchOptions): Performs a search with the specified SearchOptions on the scope of the body object. The search results are a collection of range objects.\nbody.select(selectionMode): Selects the body and navigates the Word UI to it.\nbody.select(selectionModeString): Selects the body and navigates the Word UI to it.\nbody.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\nbody.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.\nbody.onCommentAdded: Occurs when new comments are added.\nbody.onCommentChanged: Occurs when a comment or its reply is changed.\nbody.onCommentDeleted: Occurs when comments are deleted.\nbody.onCommentDeselected: Occurs when a comment is deselected.\nbody.onCommentSelected: Occurs when a comment is selected."
  },
  {
    "object": "comment",
    "apiList": "comment.authorEmail: Gets the email of the comment's author.\ncomment.authorName: Gets the name of the comment's author.\ncomment.content: Specifies the comment's content as plain text.\ncomment.contentRange: Specifies the comment's content range.\ncomment.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ncomment.creationDate: Gets the creation date of the comment.\ncomment.id: Gets the ID of the comment.\ncomment.replies: Gets the collection of reply objects associated with the comment.\ncomment.resolved: Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.\ncomment.delete(): Deletes the comment and its replies.\ncomment.getRange(): Gets the range in the main document where the comment is on.\ncomment.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncomment.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncomment.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncomment.reply(replyText): Adds a new reply to the end of the comment thread.\ncomment.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\ncomment.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object."
  },
  {
    "object": "commentCollection",
    "apiList": "body.getComments(): Gets comments associated with the body.\ncommentCollection.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ncommentCollection.items: Gets the loaded child items in this collection.\ncommentCollection.getFirst(): Gets the first comment in the collection. Throws an ItemNotFound error if this collection is empty.\ncommentCollection.getFirstOrNullObject(): Gets the first comment in the collection. If the collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ncommentCollection.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentCollection.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentCollection.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties."
  },
  {
    "object": "commentContentRange",
    "apiList": "commentContentRange.bold: Specifies a value that indicates whether the comment text is bold.\ncommentContentRange.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ncommentContentRange.hyperlink: Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.\ncommentContentRange.isEmpty: Checks whether the range length is zero.\ncommentContentRange.italic: Specifies a value that indicates whether the comment text is italicized.\ncommentContentRange.strikeThrough: Specifies a value that indicates whether the comment text has a strikethrough.\ncommentContentRange.text: Gets the text of the comment range.\ncommentContentRange.underline: Specifies a value that indicates the comment text's underline type. 'None' if the comment text isn't underlined.\ncommentContentRange.insertText(text, insertLocation): Inserts text into at the specified location. Note: For the modern comment, the content range tracked across context turns to empty if any revision to the comment is posted through the UI.\ncommentContentRange.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentContentRange.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentContentRange.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentContentRange.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\ncommentContentRange.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object."
  },
  {
    "object": "commentReply",
    "apiList": "commentReply.authorEmail: Gets the email of the comment reply's author.\ncommentReply.authorName: Gets the name of the comment reply's author.\ncommentReply.content: Specifies the comment reply's content. The string is plain text.\ncommentReply.contentRange: Specifies the commentReply's content range.\ncommentReply.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ncommentReply.creationDate: Gets the creation date of the comment reply.\ncommentReply.id: Gets the ID of the comment reply.\ncommentReply.parentComment: Gets the parent comment of this reply.\ncommentReply.delete(): Deletes the comment reply.\ncommentReply.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentReply.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentReply.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentReply.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\ncommentReply.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object."
  },
  {
    "object": "commentReplyCollection",
    "apiList": "commentReplyCollection.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ncommentReplyCollection.items: Gets the loaded child items in this collection.\ncommentReplyCollection.getFirst(): Gets the first comment reply in the collection. Throws an ItemNotFound error if this collection is empty.\ncommentReplyCollection.getFirstOrNullObject(): Gets the first comment reply in the collection. If the collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ncommentReplyCollection.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentReplyCollection.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncommentReplyCollection.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties."
  },
  {
    "object": "contentControl",
    "apiList": "contentControl.appearance: Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.\ncontentControl.cannotDelete: Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.\ncontentControl.cannotEdit: Specifies a value that indicates whether the user can edit the contents of the content control.\ncontentControl.checkboxContentControl: Specifies the checkbox-related data if the content control's type is 'CheckBox'. It's null otherwise.\ncontentControl.color: Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.\ncontentControl.contentControls: Gets the collection of content control objects in the content control.\ncontentControl.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ncontentControl.endnotes: Gets the collection of endnotes in the content control.\ncontentControl.fields: Gets the collection of field objects in the content control.\ncontentControl.font: Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.\ncontentControl.footnotes: Gets the collection of footnotes in the content control.\ncontentControl.id: Gets an integer that represents the content control identifier.\ncontentControl.inlinePictures: Gets the collection of InlinePicture objects in the content control. The collection doesn't include floating images.\ncontentControl.lists: Gets the collection of list objects in the content control.\ncontentControl.paragraphs: Gets the collection of paragraph objects in the content control.\ncontentControl.parentBody: Gets the parent body of the content control.\ncontentControl.parentContentControl: Gets the content control that contains the content control. Throws an ItemNotFound error if there isn't a parent content control.\ncontentControl.parentContentControlOrNullObject: Gets the content control that contains the content control. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ncontentControl.parentTable: Gets the table that contains the content control. Throws an ItemNotFound error if it isn't contained in a table.\ncontentControl.parentTableCell: Gets the table cell that contains the content control. Throws an ItemNotFound error if it isn't contained in a table cell.\ncontentControl.parentTableCellOrNullObject: Gets the table cell that contains the content control. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ncontentControl.parentTableOrNullObject: Gets the table that contains the content control. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ncontentControl.placeholderText: Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty. Note: The set operation for this property isn't supported in Word on the web.\ncontentControl.removeWhenEdited: Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.\ncontentControl.style: Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the 'styleBuiltIn' property.\ncontentControl.styleBuiltIn: Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the 'style' property.\ncontentControl.subtype: Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.\ncontentControl.tables: Gets the collection of table objects in the content control.\ncontentControl.tag: Specifies a tag to identify a content control.\ncontentControl.text: Gets the text of the content control.\ncontentControl.title: Specifies the title for a content control.\ncontentControl.type: Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.\ncontentControl.clear(): Clears the contents of the content control. The user can perform the undo operation on the cleared content.\ncontentControl.delete(keepContent): Deletes the content control and its content. If keepContent is set to true, the content isn't deleted.\ncontentControl.getComments(): Gets comments associated with the content control.\ncontentControl.getContentControls(options): Gets the currently supported child content controls in this content control.\ncontentControl.getHtml(): Gets an HTML representation of the content control object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use ContentControl.getOoxml() and convert the returned XML to HTML.\ncontentControl.getOoxml(): Gets the Office Open XML (OOXML) representation of the content control object.\ncontentControl.getRange(rangeLocation): Gets the whole content control, or the starting or ending point of the content control, as a range.\ncontentControl.getReviewedText(changeTrackingVersion): Gets reviewed text based on ChangeTrackingVersion selection.\ncontentControl.getReviewedText(changeTrackingVersionString): Gets reviewed text based on ChangeTrackingVersion selection.\ncontentControl.getTextRanges(endingMarks, trimSpacing): Gets the text ranges in the content control by using punctuation marks and/or other ending marks.\ncontentControl.getTrackedChanges(): Gets the collection of the TrackedChange objects in the content control.\ncontentControl.insertBreak(breakType, insertLocation): Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.\ncontentControl.insertFileFromBase64(base64File, insertLocation): Inserts a document into the content control at the specified location.\ncontentControl.insertHtml(html, insertLocation): Inserts HTML into the content control at the specified location.\ncontentControl.insertInlinePictureFromBase64(base64EncodedImage, insertLocation): Inserts an inline picture into the content control at the specified location.\ncontentControl.insertOoxml(ooxml, insertLocation): Inserts OOXML into the content control at the specified location.\ncontentControl.insertParagraph(paragraphText, insertLocation): Inserts a paragraph at the specified location.\ncontentControl.insertTable(rowCount, columnCount, insertLocation, values): Inserts a table with the specified number of rows and columns into, or next to, a content control.\ncontentControl.insertText(text, insertLocation): Inserts text into the content control at the specified location.\ncontentControl.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncontentControl.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncontentControl.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncontentControl.search(searchText, searchOptions): Performs a search with the specified SearchOptions on the scope of the content control object. The search results are a collection of range objects.\ncontentControl.select(selectionMode): Selects the content control. This causes Word to scroll to the selection.\ncontentControl.select(selectionModeString): Selects the content control. This causes Word to scroll to the selection.\ncontentControl.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\ncontentControl.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.\ncontentControl.split(delimiters, multiParagraphs, trimDelimiters, trimSpacing): Splits the content control into child ranges by using delimiters."
  },
  {
    "object": "contentControlCollection",
    "apiList": "contentControlCollection.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ncontentControlCollection.items: Gets the loaded child items in this collection.\ncontentControlCollection.getByChangeTrackingStates(changeTrackingStates): Gets the content controls that have the specified tracking state.\ncontentControlCollection.getById(id): Gets a content control by its identifier. Throws an ItemNotFound error if there isn't a content control with the identifier in this collection.\ncontentControlCollection.getByIdOrNullObject(id): Gets a content control by its identifier. If there isn't a content control with the identifier in this collection, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ncontentControlCollection.getByTag(tag): Gets the content controls that have the specified tag.\ncontentControlCollection.getByTitle(title): Gets the content controls that have the specified title.\ncontentControlCollection.getByTypes(types): Gets the content controls that have the specified types.\ncontentControlCollection.getFirst(): Gets the first content control in this collection. Throws an ItemNotFound error if this collection is empty.\ncontentControlCollection.getFirstOrNullObject(): Gets the first content control in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ncontentControlCollection.getItem(id): Gets a content control by its ID.\ncontentControlCollection.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncontentControlCollection.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ncontentControlCollection.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties."
  },
  {
    "object": "contentControlOptions",
    "apiList": "contentControlOptions.types: An array of content control types, item must be 'RichText', 'PlainText', or 'CheckBox'."
  },
  {
    "object": "document",
    "apiList": "document.body: Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.\ndocument.changeTrackingMode: Specifies the ChangeTracking mode.\ndocument.contentControls: Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.\ndocument.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ndocument.customXmlParts: Gets the custom XML parts in the document.\ndocument.properties: Gets the properties of the document.\ndocument.saved: Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.\ndocument.sections: Gets the collection of section objects in the document.\ndocument.settings: Gets the add-in's settings in the document.\ndocument.addStyle(name, type): Adds a style into the document by name and type.\ndocument.addStyle(name, typeString): Adds a style into the document by name and type.\ndocument.close(closeBehavior): Closes the current document.\ndocument.close(closeBehaviorString): Closes the current document.\ndocument.compare(filePath, documentCompareOptions): Displays revision marks that indicate where the specified document differs from another document.\ndocument.deleteBookmark(name): Deletes a bookmark, if it exists, from the document.\ndocument.getAnnotationById(id): Gets the annotation by ID. Throws an ItemNotFound error if annotation isn't found.\ndocument.getBookmarkRange(name): Gets a bookmark's range. Throws an ItemNotFound error if the bookmark doesn't exist.\ndocument.getBookmarkRangeOrNullObject(name): Gets a bookmark's range. If the bookmark doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ndocument.getContentControls(options): Gets the currently supported content controls in the document.\ndocument.getEndnoteBody(): Gets the document's endnotes in a single body.\ndocument.getFootnoteBody(): Gets the document's footnotes in a single body.\ndocument.getParagraphByUniqueLocalId(id): Gets the paragraph by its unique local ID. Throws an ItemNotFound error if the collection is empty.\ndocument.getSelection(): Gets the current selection of the document. Multiple selections aren't supported.\ndocument.getStyles(): Gets a StyleCollection object that represents the whole style set of the document.\ndocument.importStylesFromJson(stylesJson): Import styles from a JSON-formatted string.\ndocument.insertFileFromBase64(base64File, insertLocation, insertFileOptions): Inserts a document into the target document at a specific location with additional properties. Headers, footers, watermarks, and other section properties are copied by default.\ndocument.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ndocument.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ndocument.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ndocument.save(saveBehavior, fileName): Saves the document.\ndocument.save(saveBehaviorString, fileName): Saves the document.\ndocument.search(searchText, searchOptions): Performs a search with the specified search options on the scope of the whole document. The search results are a collection of range objects.\ndocument.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\ndocument.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object."
  },
  {
    "object": "documentCreated",
    "apiList": "documentCreated.body: Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.\ndocumentCreated.contentControls: Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.\ndocumentCreated.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\ndocumentCreated.customXmlParts: Gets the custom XML parts in the document.\ndocumentCreated.properties: Gets the properties of the document.\ndocumentCreated.saved: Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.\ndocumentCreated.sections: Gets the collection of section objects in the document.\ndocumentCreated.settings: Gets the add-in's settings in the document.\ndocumentCreated.addStyle(name, type): Adds a style into the document by name and type.\ndocumentCreated.addStyle(name, typeString): Adds a style into the document by name and type.\ndocumentCreated.deleteBookmark(name): Deletes a bookmark, if it exists, from the document.\ndocumentCreated.getBookmarkRange(name): Gets a bookmark's range. Throws an ItemNotFound error if the bookmark doesn't exist.\ndocumentCreated.getBookmarkRangeOrNullObject(name): Gets a bookmark's range. If the bookmark doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\ndocumentCreated.getContentControls(options): Gets the currently supported content controls in the document.\ndocumentCreated.getStyles(): Gets a StyleCollection object that represents the whole style set of the document.\ndocumentCreated.insertFileFromBase64(base64File, insertLocation, insertFileOptions): Inserts a document into the target document at a specific location with additional properties. Headers, footers, watermarks, and other section properties are copied by default.\ndocumentCreated.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ndocumentCreated.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ndocumentCreated.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\ndocumentCreated.open(): Opens the document.\ndocumentCreated.save(saveBehavior, fileName): Saves the document.\ndocumentCreated.save(saveBehaviorString, fileName): Saves the document.\ndocumentCreated.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\ndocumentCreated.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object."
  },
  {
    "object": "paragraph",
    "apiList": "paragraph.alignment: Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.\nparagraph.contentControls: Gets the collection of content control objects in the paragraph.\nparagraph.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\nparagraph.endnotes: Gets the collection of endnotes in the paragraph.\nparagraph.fields: Gets the collection of fields in the paragraph.\nparagraph.firstLineIndent: Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.\nparagraph.font: Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.\nparagraph.footnotes: Gets the collection of footnotes in the paragraph.\nparagraph.inlinePictures: Gets the collection of InlinePicture objects in the paragraph. The collection doesn't include floating images.\nparagraph.isLastParagraph: Indicates the paragraph is the last one inside its parent body.\nparagraph.isListItem: Checks whether the paragraph is a list item.\nparagraph.leftIndent: Specifies the left indent value, in points, for the paragraph.\nparagraph.lineSpacing: Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.\nparagraph.lineUnitAfter: Specifies the amount of spacing, in grid lines, after the paragraph.\nparagraph.lineUnitBefore: Specifies the amount of spacing, in grid lines, before the paragraph.\nparagraph.list: Gets the List to which this paragraph belongs. Throws an ItemNotFound error if the paragraph isn't in a list.\nparagraph.listItem: Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.\nparagraph.listItemOrNullObject: Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraph.listOrNullObject: Gets the List to which this paragraph belongs. If the paragraph isn't in a list, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraph.outlineLevel: Specifies the outline level for the paragraph.\nparagraph.parentBody: Gets the parent body of the paragraph.\nparagraph.parentContentControl: Gets the content control that contains the paragraph. Throws an ItemNotFound error if there isn't a parent content control.\nparagraph.parentContentControlOrNullObject: Gets the content control that contains the paragraph. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraph.parentTable: Gets the table that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table.\nparagraph.parentTableCell: Gets the table cell that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table cell.\nparagraph.parentTableCellOrNullObject: Gets the table cell that contains the paragraph. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraph.parentTableOrNullObject: Gets the table that contains the paragraph. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraph.rightIndent: Specifies the right indent value, in points, for the paragraph.\nparagraph.spaceAfter: Specifies the spacing, in points, after the paragraph.\nparagraph.spaceBefore: Specifies the spacing, in points, before the paragraph.\nparagraph.style: Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the 'styleBuiltIn' property.\nparagraph.styleBuiltIn: Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the 'style' property.\nparagraph.tableNestingLevel: Gets the level of the paragraph's table. It returns 0 if the paragraph isn't in a table.\nparagraph.text: Gets the text of the paragraph.\nparagraph.uniqueLocalId: Gets a string that represents the paragraph identifier in the current session. ID is in standard 8-4-4-4-12 GUID format without curly braces and differs across sessions and coauthors.\nparagraph.attachToList(listId, level): Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.\nparagraph.clear(): Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.\nparagraph.delete(): Deletes the paragraph and its content from the document.\nparagraph.detachFromList(): Moves this paragraph out of its list, if the paragraph is a list item.\nparagraph.getAnnotations(): Gets annotations set on this Paragraph object.\nparagraph.getComments(): Gets comments associated with the paragraph.\nparagraph.getContentControls(options): Gets the currently supported content controls in the paragraph.\nparagraph.getHtml(): Gets an HTML representation of the paragraph object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use Paragraph.getOoxml() and convert the returned XML to HTML.\nparagraph.getNext(): Gets the next paragraph. Throws an ItemNotFound error if the paragraph is the last one.\nparagraph.getNextOrNullObject(): Gets the next paragraph. If the paragraph is the last one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraph.getOoxml(): Gets the Office Open XML (OOXML) representation of the paragraph object.\nparagraph.getPrevious(): Gets the previous paragraph. Throws an ItemNotFound error if the paragraph is the first one.\nparagraph.getPreviousOrNullObject(): Gets the previous paragraph. If the paragraph is the first one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraph.getRange(rangeLocation): Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.\nparagraph.getReviewedText(changeTrackingVersion): Gets reviewed text based on ChangeTrackingVersion selection.\nparagraph.getReviewedText(changeTrackingVersionString): Gets reviewed text based on ChangeTrackingVersion selection.\nparagraph.getText(options): Returns the text of the paragraph. This excludes equations, graphics (e.g., images, videos, drawings), and special characters that mark various content (e.g., for content controls, fields, comments, footnotes, endnotes). By default, hidden text and text marked as deleted are excluded.\nparagraph.getTextRanges(endingMarks, trimSpacing): Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.\nparagraph.getTrackedChanges(): Gets the collection of the TrackedChange objects in the paragraph.\nparagraph.insertAnnotations(annotations): Inserts annotations on this Paragraph object.\nparagraph.insertBreak(breakType, insertLocation): Inserts a break at the specified location in the main document.\nparagraph.insertContentControl(contentControlType): Wraps the Paragraph object with a content control.\nparagraph.insertFileFromBase64(base64File, insertLocation): Inserts a document into the paragraph at the specified location.\nparagraph.insertHtml(html, insertLocation): Inserts HTML into the paragraph at the specified location.\nparagraph.insertInlinePictureFromBase64(base64EncodedImage, insertLocation): Inserts a picture into the paragraph at the specified location.\nparagraph.insertOoxml(ooxml, insertLocation): Inserts OOXML into the paragraph at the specified location.\nparagraph.insertParagraph(paragraphText, insertLocation): Inserts a paragraph at the specified location.\nparagraph.insertTable(rowCount, columnCount, insertLocation, values): Inserts a table with the specified number of rows and columns.\nparagraph.insertText(text, insertLocation): Inserts text into the paragraph at the specified location.\nparagraph.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nparagraph.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nparagraph.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nparagraph.search(searchText, searchOptions): Performs a search with the specified SearchOptions on the scope of the paragraph object. The search results are a collection of range objects.\nparagraph.select(selectionMode): Selects and navigates the Word UI to the paragraph.\nparagraph.select(selectionModeString): Selects and navigates the Word UI to the paragraph.\nparagraph.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\nparagraph.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.\nparagraph.split(delimiters, trimDelimiters, trimSpacing): Splits the paragraph into child ranges by using delimiters.\nparagraph.startNewList(): Starts a new list with this paragraph. Fails if the paragraph is already a list item."
  },
  {
    "object": "paragraphCollection",
    "apiList": "paragraphCollection.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\nparagraphCollection.items: Gets the loaded child items in this collection.\nparagraphCollection.getFirst(): Gets the first paragraph in this collection. Throws an ItemNotFound error if the collection is empty.\nparagraphCollection.getFirstOrNullObject(): Gets the first paragraph in this collection. If the collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraphCollection.getLast(): Gets the last paragraph in this collection. Throws an ItemNotFound error if the collection is empty.\nparagraphCollection.getLastOrNullObject(): Gets the last paragraph in this collection. If the collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nparagraphCollection.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nparagraphCollection.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nparagraphCollection.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties."
  },
  {
    "object": "paragraphCollection",
    "apiList": "range.contentControls: Gets the collection of content control objects in the range.\nrange.context: The request context associated with the object. This connects the add-in's process to the Office host application's process.\nrange.endnotes: Gets the collection of endnotes in the range.\nrange.fields: Gets the collection of field objects in the range.\nrange.font: Gets the text format of the range. Use this to get and set font name, size, color, and other properties.\nrange.footnotes: Gets the collection of footnotes in the range.\nrange.hyperlink: Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.\nrange.inlinePictures: Gets the collection of inline picture objects in the range.\nrange.isEmpty: Checks whether the range length is zero.\nrange.lists: Gets the collection of list objects in the range.\nrange.paragraphs: Gets the collection of paragraph objects in the range.\nrange.parentBody: Gets the parent body of the range.\nrange.parentContentControl: Gets the currently supported content control that contains the range. Throws an ItemNotFound error if there isn't a parent content control.\nrange.parentContentControlOrNullObject: Gets the currently supported content control that contains the range. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nrange.parentTable: Gets the table that contains the range. Throws an ItemNotFound error if it isn't contained in a table.\nrange.parentTableCell: Gets the table cell that contains the range. Throws an ItemNotFound error if it isn't contained in a table cell.\nrange.parentTableCellOrNullObject: Gets the table cell that contains the range. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nrange.parentTableOrNullObject: Gets the table that contains the range. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nrange.style: Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the 'styleBuiltIn' property.\nrange.styleBuiltIn: Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the 'style' property.\nrange.tables: Gets the collection of table objects in the range.\nrange.text: Gets the text of the range.\nrange.clear(): Clears the contents of the range object. The user can perform the undo operation on the cleared content.\nrange.compareLocationWith(range): Compares this range's location with another range's location.\nrange.delete(): Deletes the range and its content from the document.\nrange.expandTo(range): Returns a new range that extends from this range in either direction to cover another range. This range isn't changed. Throws an ItemNotFound error if the two ranges do not have a union.\nrange.expandToOrNullObject(range): Returns a new range that extends from this range in either direction to cover another range. This range isn't changed. If the two ranges don't have a union, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nrange.getBookmarks(includeHidden, includeAdjacent): Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with the underscore character.\nrange.getComments(): Gets comments associated with the range.\nrange.getContentControls(options): Gets the currently supported content controls in the range.\nrange.getHtml(): Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use Range.getOoxml() and convert the returned XML to HTML.\nrange.getHyperlinkRanges(): Gets hyperlink child ranges within the range.\nrange.getNextTextRange(endingMarks, trimSpacing): Gets the next text range by using punctuation marks and/or other ending marks. Throws an ItemNotFound error if this text range is the last one.\nrange.getNextTextRangeOrNullObject(endingMarks, trimSpacing): Gets the next text range by using punctuation marks and/or other ending marks. If this text range is the last one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nrange.getOoxml(): Gets the OOXML representation of the range object.\nrange.getRange(rangeLocation): Clones the range, or gets the starting or ending point of the range as a new range.\nrange.getReviewedText(changeTrackingVersion): Gets reviewed text based on ChangeTrackingVersion selection.\nrange.getReviewedText(changeTrackingVersionString): Gets reviewed text based on ChangeTrackingVersion selection.\nrange.getTextRanges(endingMarks, trimSpacing): Gets the text child ranges in the range by using punctuation marks and/or other ending marks.\nrange.getTrackedChanges(): Gets the collection of the TrackedChange objects in the range.\nrange.highlight(): Highlights the range temporarily without changing document content. To highlight the text permanently, set the range's Font.HighlightColor.\nrange.insertBookmark(name): Inserts a bookmark on the range. If a bookmark of the same name exists somewhere, it is deleted first.\nrange.insertBreak(breakType, insertLocation): Inserts a break at the specified location in the main document.\nrange.insertComment(commentText): Insert a comment on the range.\nrange.insertContentControl(contentControlType): Wraps the Range object with a content control.\nrange.insertEndnote(insertText): Inserts an endnote. The endnote reference is placed after the range.\nrange.insertField(insertLocation, fieldType, text, removeFormatting): Inserts a field at the specified location.\nrange.insertField(insertLocation, fieldTypeString, text, removeFormatting): Inserts a field at the specified location.\nrange.insertFileFromBase64(base64File, insertLocation): Inserts a document at the specified location.\nrange.insertFootnote(insertText): Inserts a footnote. The footnote reference is placed after the range.\nrange.insertHtml(html, insertLocation): Inserts HTML at the specified location.\nrange.insertInlinePictureFromBase64(base64EncodedImage, insertLocation): Inserts a picture at the specified location.\nrange.insertOoxml(ooxml, insertLocation): Inserts OOXML at the specified location.\nrange.insertParagraph(paragraphText, insertLocation): Inserts a paragraph at the specified location.\nrange.insertTable(rowCount, columnCount, insertLocation, values): Inserts a table with the specified number of rows and columns.\nrange.insertText(text, insertLocation): Inserts text at the specified location.\nrange.intersectWith(range): Returns a new range as the intersection of this range with another range. This range isn't changed. Throws an ItemNotFound error if the two ranges aren't overlapped or adjacent.\nrange.intersectWithOrNullObject(range): Returns a new range as the intersection of this range with another range. This range isn't changed. If the two ranges aren't overlapped or adjacent, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.\nrange.load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nrange.load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nrange.load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.\nrange.removeHighlight(): Removes the highlight added by the Highlight function if any.\nrange.search(searchText, searchOptions): Performs a search with the specified SearchOptions on the scope of the range object. The search results are a collection of range objects.\nrange.select(selectionMode): Selects and navigates the Word UI to the range.\nrange.select(selectionModeString): Selects and navigates the Word UI to the range.\nrange.set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.\nrange.set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.\nrange.split(delimiters, multiParagraphs, trimDelimiters, trimSpacing): Splits the range into child ranges by using delimiters."
  },
  {
    "object": "Excel.Chart",
    "apiList": ` worksheet.charts: Returns a collection of charts that are part of the worksheet.
    chart.axes: Represents chart axes.
    chart.categoryLabelLevel: Specifies a chart category label level enumeration constant, referring to the level of the source category labels.
    chart.chartType: Specifies the type of the chart. See Excel.ChartType for details.
    chart.dataLabels: Represents the data labels on the chart.
    chart.displayBlanksAs: Specifies the way that blank cells are plotted on a chart.
    chart.format: Encapsulates the format properties for the chart area.
    chart.height: Specifies the height, in points, of the chart object.
    chart.id: The unique ID of chart.
    chart.left: The distance, in points, from the left side of the chart to the worksheet origin.
    chart.legend: Represents the legend for the chart.
    chart.name: Specifies the name of a chart object.
    chart.pivotOptions: Encapsulates the options for a pivot chart.
    chart.plotArea: Represents the plot area for the chart.
    chart.plotBy: Specifies the way columns or rows are used as data series on the chart.
    chart.plotVisibleOnly: True if only visible cells are plotted. False if both visible and hidden cells are plotted.
    chart.series: Represents either a single series or collection of series in the chart.
    chart.seriesNameLevel: Specifies a chart series name level enumeration constant, referring to the level of the source series names.
    chart.showAllFieldButtons: Specifies whether to display all field buttons on a PivotChart.
    chart.showDataLabelsOverMaximum: Specifies whether to show the data labels when the value is greater than the maximum value on the value axis.
    chart.style: Specifies the chart style for the chart.
    chart.title: Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.
    chart.top: Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
    chart.width: Specifies the width, in points, of the chart object.
    chart.worksheet: The worksheet containing the current chart.
    chart.activate(): Activates the chart in the Excel UI.
    chart.delete(): Deletes the chart object.
    chart.getDataTable(): Gets the data table on the chart.
    chart.getDataTableOrNullObject(): Gets the data table on the chart, returning an object if the chart doesn't allow a data table.
    chart.getImage(width, height, fittingMode): Renders the chart as a base64-encoded image.
    chart.load(options): Queues up a command to load specified properties of the object.
    chart.set(properties, options): Sets multiple properties of an object at the same time.
    chart.setData(sourceData, seriesBy): Resets the source data for the chart.
    chart.setPosition(startCell, endCell): Positions the chart relative to cells on the worksheet.`
  },
  {
    "object": "Excel.ChartCollection",
    "apiList": ` workbook.worksheets: Represents a collection of worksheets associated with the workbook.
    worksheets.getActiveWorksheet(): Gets the currently active worksheet in the workbook.
    worksheet.charts: Returns a collection of charts that are part of the worksheet.
    chatCollection.count: Returns the number of charts in the worksheet.
    chatCollection.items: Gets the loaded child items in this collection.
    chatCollection.add(type: Excel.ChartType, sourceData: Excel.Range, seriesBy: Excel.ChartSeriesBy): Creates a new chart.
    chatCollection.getCount(): Returns the number of charts in the worksheet.
    chatCollection.getItem(name): Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
    chatCollection.getItemAt(index): Gets a chart based on its position in the collection.
    chatCollection.getItemOrNullObject(name): Gets a chart using its name; returns an object with isNullObject true if the chart doesn't exist.
    chatCollection.load(options): Queues up a command to load specified properties of the object. Requires context.sync() before reading properties.`
  },
  {
    "object": "Excel.ChartType",
    "apiList": ` _3DArea
    _3DAreaStacked
    _3DAreaStacked100
    _3DBarClustered
    _3DBarStacked
    _3DBarStacked100
    _3DColumn
    _3DColumnClustered
    _3DColumnStacked
    _3DColumnStacked100
    _3DLine
    _3DPie
    _3DPieExploded
    area
    areaStacked
    areaStacked100
    barClustered
    barOfPie
    barStacked
    barStacked100
    boxwhisker
    bubble
    bubble3DEffect
    columnClustered
    columnStacked
    columnStacked100
    coneBarClustered
    coneBarStacked
    coneBarStacked100
    coneCol
    coneColClustered
    coneColStacked
    coneColStacked100
    cylinderBarClustered
    cylinderBarStacked
    cylinderBarStacked100
    cylinderCol
    cylinderColClustered
    cylinderColStacked
    cylinderColStacked100
    doughnut
    doughnutExploded
    funnel
    histogram
    invalid
    line
    lineMarkers
    lineMarkersStacked
    lineMarkersStacked100
    lineStacked
    lineStacked100
    pareto
    pie
    pieExploded
    pieOfPie
    pyramidBarClustered
    pyramidBarStacked
    pyramidBarStacked100
    pyramidCol
    pyramidColClustered
    pyramidColStacked
    pyramidColStacked100
    radar
    radarFilled
    radarMarkers
    regionMap
    stockHLC
    stockOHLC
    stockVHLC
    stockVOHLC
    sunburst
    surface
    surfaceTopView
    surfaceTopViewWireframe
    surfaceWireframe
    treemap
    waterfall
    xyscatter
    xyscatterLines
    xyscatterLinesNoMarkers
    xyscatterSmooth
    xyscatterSmoothNoMarkers
    `
  }
];

function correctEnumSpelling(enumString: string): string {

  const regex = /Excel.ChartType.([\s\S]*?),/g;
  const matches = [...enumString.matchAll(regex)];
  const codeMath = matches.map((match) => match[1]).join('\n');
  const lowerCaseStarted = codeMath.charAt(0).toLowerCase() + codeMath.slice(1);

  return enumString.split(codeMath).join(lowerCaseStarted);
}
