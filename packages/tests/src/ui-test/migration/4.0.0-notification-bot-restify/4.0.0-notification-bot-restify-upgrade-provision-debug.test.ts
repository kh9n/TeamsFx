// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { MigrationTestContext } from "../migrationContext";
import {
  Timeout,
  Capability,
  Trigger,
  Notification,
} from "../../../utils/constants";
import { it } from "../../../utils/it";
import { Env } from "../../../utils/env";
import {
  validateNotificationBot,
  initPage,
} from "../../../utils/playwrightOperation";
import { CliHelper } from "../../cliHelper";
import {
  validateNotification,
  upgradeByTreeView,
  validateUpgrade,
} from "../../../utils/vscodeOperation";
import {
  CLIVersionCheck,
  getBotSiteEndpoint,
  updateDeverloperInManifestFile,
} from "../../../utils/commonUtils";
import { updatePakcageJson } from "./helper";
import path from "path";
import { runDeploy, runProvision } from "../../remotedebug/remotedebugContext";

describe("Migration Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let mirgationDebugTestContext: MigrationTestContext;
  CliHelper.setV3Enable();

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);

    mirgationDebugTestContext = new MigrationTestContext(
      Capability.Notification,
      "javascript",
      Trigger.Restify
    );
    await mirgationDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await mirgationDebugTestContext.after(false, true, "dev");
  });

  it(
    "[auto] V4.0.0 notification bot template upgrade test - js",
    {
      testPlanCaseId: 17431842,
      author: "frankqian@microsoft.com",
    },
    async () => {
      // create v2 project using CLI
      await mirgationDebugTestContext.createProjectCLI(false);

      // update package.json in bot folder
      await updatePakcageJson(
        path.join(mirgationDebugTestContext.projectPath, "bot", "package.json")
      );

      // verify popup
      await validateNotification(Notification.Upgrade);

      // upgrade
      await upgradeByTreeView();
      // verify upgrade
      await validateUpgrade();
      // enable cli v3
      CliHelper.setV3Enable();

      // install test cil in project
      await CliHelper.installCLI(
        Env.TARGET_CLI,
        false,
        mirgationDebugTestContext.testRootFolder
      );
      // enable cli v3
      CliHelper.setV3Enable();

      await updateDeverloperInManifestFile(
        mirgationDebugTestContext.projectPath
      );

      // v3 provision
      await runProvision(mirgationDebugTestContext.appName);
      await runDeploy(Timeout.botDeploy * 2);

      const teamsAppId = await mirgationDebugTestContext.getTeamsAppId("dev");

      // UI verify
      const page = await initPage(
        mirgationDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password
      );
      const funcEndpoint = await getBotSiteEndpoint(
        mirgationDebugTestContext.projectPath,
        "dev"
      );
      await validateNotificationBot(page, funcEndpoint + "/api/notification");
    }
  );
});
