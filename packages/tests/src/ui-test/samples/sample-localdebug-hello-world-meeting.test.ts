// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { initTeamsPage } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";
import { SampledebugContext } from "./sampledebugContext";

class MyFirstMettingTestCase extends CaseFactory {}

new MyFirstMettingTestCase(
  TemplateProject.MyFirstMetting,
  9958524,
  "v-ivanchen@microsoft.com",
  "local",
  [LocalDebugTaskLabel.StartFrontend],
  {
    teamsAppName: "hello-world-in-meetinglocal",
    type: "meeting",
    skipValidation: true,
    debug: "cli",
  }
).test();
