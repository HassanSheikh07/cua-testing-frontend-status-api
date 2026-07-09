/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { ConnectionSettings } from '@microsoft/agents-copilotstudio-client'

window.localStorage.debug = 'copilot-studio:*'

export class SampleConnectionSettings extends ConnectionSettings {
  constructor () {
    super({
      // Leave these empty because we are using directConnectUrl
      environmentId: '',
      schemaName: '',

      // Paste the Connection string copied from:
      // Channels → Web app → Microsoft 365 Agents SDK → Connection string
      directConnectUrl: 'https://7236333b737ce8549c86767a9b3e1c.02.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/cr720_Agent1TestScript/conversations?api-version=2022-03-01-preview',

      // Leave default/empty for normal commercial cloud
      cloud: '',
      customPowerPlatformCloud: '',

      // Published agent
      copilotAgentType: 'Published',

      useExperimentalEndpoint: false
    })

    // App Registration Client ID
    this.appClientId = '5e69a3e2-b3bc-4ae2-8df9-557556be7cd5'

    // Tenant ID where the Copilot Studio agent exists
    this.tenantId = 'edda99bb-bab6-4c4c-8aa1-4b99e8e09c1b'

    // Usually keep this empty, or set it explicitly
    this.authority = 'https://login.microsoftonline.com'
  }
}