import { BaseWebQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IAceDynamicFlowCardAdaptiveCardExtensionProps,
  IAceDynamicFlowCardAdaptiveCardExtensionState
} from '../AceDynamicFlowCardAdaptiveCardExtension';
import { HttpClientResponse } from '@microsoft/sp-http'; // Removed AadHttpClient since we're not using it
import { AadTokenProvider } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';
import { jwtDecode } from 'jwt-decode'; // Changed from default import to named import

interface IDecodedToken {
  tid: string; // tenant id
  oid: string; // object id
  upn: string; // user principal name
  name: string; // display name
  [key: string]: any;
}

export class QuickView extends BaseWebQuickView<
  IAceDynamicFlowCardAdaptiveCardExtensionProps,
  IAceDynamicFlowCardAdaptiveCardExtensionState
> {

  public render(): void {
    console.log('🎨 QuickView render() called');
    console.log('🔍 State in render:', this.state);
    console.log('🔍 Properties in render:', this.properties);

    // Check if we need to call the flow
    if (this.state.isLoading && !this.state.error && !this.state.htmlContent) {
      console.log('⚡ Initial render - calling flow immediately');
      this.callPowerAutomateFlow();
    }

    this.renderContent();
  }

  private renderContent(): void {
    console.log('🖼️ renderContent called, state:', {
      isLoading: this.state.isLoading,
      hasError: !!this.state.error,
      hasContent: !!this.state.htmlContent,
      contentLength: this.state.htmlContent?.length || 0
    });

    if (!this.state) {
      console.error('💥 State is undefined');
      this.domElement.innerHTML = `
        <div style="padding: 20px; border: 1px solid #dc3545; background-color: #f8d7da;">
          <h3>State Error</h3>
          <p>Component state is undefined. Check console for details.</p>
        </div>`;
      return;
    }

    if (this.state.isLoading) {
      console.log('⏳ Rendering loading state');
      this.domElement.innerHTML = `
        <div style="padding: 20px; text-align: center; border: 2px solid #007acc; border-radius: 8px; background-color: #f0f8ff;">
          <div style="margin-bottom: 10px; font-size: 24px;">⏳</div>
          <p style="margin: 0; font-weight: bold;">Loading Power Automate Flow...</p>
          <small style="color: #666;">Using PnP Bearer token approach</small>
          <div style="margin-top: 15px;">
            <div style="display: inline-block; width: 20px; height: 20px; border: 3px solid #007acc; border-top-color: transparent; border-radius: 50%; animation: spin 1s linear infinite;"></div>
          </div>
          <div style="margin-top: 10px;">
            <small>Properties: ${JSON.stringify(this.properties)}</small>
          </div>
        </div>
        <style>
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>`;
      return;
    }

    if (this.state.error) {
      console.log('❌ Rendering error state:', this.state.error);
      this.domElement.innerHTML = this.state.htmlContent || `
        <div style="padding: 20px; border: 1px solid #dc3545; background-color: #f8d7da;">
          <h3>Error</h3>
          <p>${escape(this.state.error)}</p>
        </div>`;
      return;
    }

    if (this.state.htmlContent && this.state.htmlContent.trim().length > 0) {
      console.log('✅ Rendering Flow HTML content');
      console.log('✅ Content to render:', this.state.htmlContent.substring(0, 200) + '...');
      
      // Wrap the flow content in a container for better styling
      this.domElement.innerHTML = `
        <div style="padding: 10px; background-color: #fff; border-radius: 4px;">
          <div style="margin-bottom: 10px; font-size: 12px; color: #666; border-bottom: 1px solid #eee; padding-bottom: 5px;">
            ✅ Content from Power Automate Flow
          </div>
          ${this.state.htmlContent}
        </div>`;
      
      console.log('✅ HTML content successfully set to domElement');
    } else {
      console.log('⚠️ Rendering fallback - no content available');
      this.domElement.innerHTML = `
        <div style="padding: 20px; border: 1px solid #ccc; background-color: #f8f9fa;">
          <h3>No Content</h3>
          <p>No content available from Power Automate flow.</p>
          <p><strong>Properties:</strong></p>
          <pre>${JSON.stringify(this.properties, null, 2)}</pre>
          <p><strong>State:</strong></p>
          <pre>${JSON.stringify(this.state, null, 2)}</pre>
        </div>`;
    }
  }

  private async callPowerAutomateFlow(): Promise<void> {
    console.log('🚀 callPowerAutomateFlow started');
    
    const flowUrl = this.properties.flowUrl || this.properties.powerAutomateUrl;
    console.log('🔗 Flow URL from properties:', flowUrl);
    console.log('🔗 Prompt from properties:', this.properties.prompt);
    console.log('📊 All properties:', JSON.stringify(this.properties, null, 2));
    
    if (!flowUrl || flowUrl.trim() === '') {
      console.warn('⚠️ No flow URL configured');
      this.setState({
        htmlContent: `
          <div style="padding: 20px; border: 1px solid #ffa500; background-color: #fff3cd;">
            <h3>Configuration Required</h3>
            <p>Please configure the Power Automate Flow URL in the web part properties.</p>
            <p><strong>Current Properties:</strong></p>
            <pre>${JSON.stringify(this.properties, null, 2)}</pre>
            <p><strong>How to configure:</strong></p>
            <ol>
              <li>Click the gear icon on the ACE card</li>
              <li>Enter your Power Automate Flow HTTP trigger URL</li>
              <li>Enter a prompt or instruction for your flow (optional)</li>
              <li>Save the configuration</li>
            </ol>
            <p><strong>PnP Recommended Flow Setup:</strong></p>
            <ul>
              <li>Use "When an HTTP request is received" trigger</li>
              <li>Configure proper JSON schema for request body (include "prompt" field)</li>
              <li>Use the prompt field in your flow logic</li>
              <li>Use "Response" action to return HTML content</li>
              <li>Ensure "Microsoft Flow Service" API permissions are approved</li>
            </ul>
          </div>`,
        isLoading: false,
        error: 'No flow URL configured'
      });
      return;
    }

    // Validate URL format
    try {
      new URL(flowUrl);
      console.log('✅ URL format validation passed');
    } catch (urlError) {
      console.error('❌ Invalid URL format:', urlError);
      this.setState({
        htmlContent: `
          <div style="padding: 20px; border: 1px solid #dc3545; background-color: #f8d7da;">
            <h3>Invalid URL</h3>
            <p>The configured Flow URL is not a valid URL format.</p>
            <p><strong>URL:</strong> ${escape(flowUrl)}</p>
            <p><strong>Expected format:</strong> https://prod-xx.westus.logic.azure.com:443/workflows/...</p>
          </div>`,
        isLoading: false,
        error: 'Invalid URL format'
      });
      return;
    }

    // Call flow using PnP recommended approach with AadTokenProvider
    await this.callFlowWithPnpApproach(flowUrl);
  }

  private async callFlowWithPnpApproach(flowUrl: string): Promise<void> {
    try {
      console.log('🔐 Using PnP article approach: Bearer token with direct fetch');
      
      // Step 1: Get the token provider (as per PnP article)
      const provider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      console.log('✅ AadTokenProvider obtained');
      
      // Step 2: Get token for Flow service (as per PnP article)
      const token: string = await provider.getToken("https://service.flow.microsoft.com/");
      console.log('✅ Token obtained for Flow service');
      
      // Step 3: Decode the token to get tenant ID (as per PnP article)
      const decodedToken: IDecodedToken = jwtDecode(token);
      const envId: string = decodedToken.tid; // tenant ID
      
      console.log('🔍 Decoded token info:', {
        tenantId: envId,
        objectId: decodedToken.oid,
        userPrincipalName: decodedToken.upn,
        displayName: decodedToken.name,
        tokenLength: token.length,
        tokenPreview: token.substring(0, 50) + '...'
      });

      // Step 4: Prepare request payload with prompt
      const requestPayload = {
        userId: this.context.pageContext.user.loginName || 'unknown',
        userEmail: this.context.pageContext.user.email || '',
        displayName: this.context.pageContext.user.displayName || '',
        siteUrl: this.context.pageContext.web.absoluteUrl || '',
        timestamp: new Date().toISOString(),
        requestId: Math.random().toString(36).substr(2, 9),
        tenantId: envId,
        prompt: this.properties.prompt || '' // Include prompt from properties
      };

      console.log('📤 Request payload:', JSON.stringify(requestPayload, null, 2));
      console.log('📡 Making authenticated fetch request with Bearer token to:', flowUrl);
      
      // Step 5: Use fetch with Bearer token (as per PnP article)
      const response = await fetch(flowUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
          'Accept': 'text/html,application/json,text/plain,*/*'
        },
        body: JSON.stringify(requestPayload)
      });

      // Step 6: Process response
      const adaptedResponse = {
        ok: response.ok,
        status: response.status,
        statusText: response.statusText,
        text: async () => await response.text(),
        url: response.url
      } as HttpClientResponse;

      await this.processFlowResponse(adaptedResponse, flowUrl, `PnP Bearer Token (Tenant: ${envId})`);

    } catch (error) {
      console.error('❌ PnP Bearer token approach failed:', error);
      console.error('❌ Error details:', {
        message: error.message,
        stack: error.stack,
        name: error.name
      });
      
      // Fallback to direct fetch for public endpoints
      try {
        console.log('🔓 Fallback: Trying direct fetch (public endpoint)');
        await this.callFlowWithDirectFetch(flowUrl);
      } catch (fetchError) {
        console.error('❌ All authentication methods failed');
        this.showPnpAuthenticationFailure(flowUrl, error.message, fetchError.message);
      }
    }
  }

  private async callFlowWithDirectFetch(flowUrl: string): Promise<void> {
    console.log('🔓 Using direct fetch (no authentication)');
    
    // Simplified payload for direct fetch too - include prompt
    const requestPayload = {
      userId: this.context.pageContext.user.loginName || 'unknown',
      userEmail: this.context.pageContext.user.email || '',
      displayName: this.context.pageContext.user.displayName || '',
      siteUrl: this.context.pageContext.web.absoluteUrl || '',
      timestamp: new Date().toISOString(),
      requestId: Math.random().toString(36).substr(2, 9),
      prompt: this.properties.prompt || '' // Include prompt from properties
    };

    const response = await fetch(flowUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'text/html,application/json,text/plain,*/*'
      },
      body: JSON.stringify(requestPayload)
    });

    const adaptedResponse = {
      ok: response.ok,
      status: response.status,
      statusText: response.statusText,
      text: async () => await response.text(),
      url: response.url
    } as HttpClientResponse;

    await this.processFlowResponse(adaptedResponse, flowUrl, 'Direct Fetch (No Authentication)');
  }

  private async processFlowResponse(response: HttpClientResponse, flowUrl: string, authMethod: string): Promise<void> {
    console.log('📥 Response received:', {
      status: response.status,
      statusText: response.statusText,
      ok: response.ok,
      authMethod: authMethod
    });

    if (response.ok) {
      const htmlContent = await response.text();
      console.log('📄 Content received (length):', htmlContent.length);
      console.log('📄 Raw content preview:', htmlContent.substring(0, 200) + '...');
      console.log('📄 Content type check:', typeof htmlContent);
      
      // Trim whitespace and check if content exists
      const trimmedContent = htmlContent.trim();
      
      if (trimmedContent.length === 0) {
        this.setState({
          htmlContent: `
            <div style="padding: 20px; border: 1px solid #ffa500; background-color: #fff3cd;">
              <h3>✅ Flow Executed Successfully</h3>
              <p>The Power Automate flow was called successfully but returned empty content.</p>
              <p><strong>Status:</strong> ${response.status} ${response.statusText}</p>
              <p><strong>Authentication Method:</strong> ${escape(authMethod)}</p>
              <p><strong>Flow URL:</strong> ${escape(flowUrl)}</p>
              <div style="margin-top: 15px; padding: 10px; background-color: #e7f3ff; border-radius: 4px;">
                <p><strong>💡 Note:</strong> Your flow executed successfully using PnP recommended authentication but didn't return any HTML content.</p>
                <p><strong>To fix this:</strong></p>
                <ol>
                  <li>Add a "Response" action at the end of your flow</li>
                  <li>Set the status code to 200</li>
                  <li>Set Content-Type header to "text/html"</li>
                  <li>Add your HTML content in the body</li>
                </ol>
              </div>
            </div>`,
          isLoading: false,
          error: ''
        });
      } else {
        console.log('✅ Setting HTML content in state:', trimmedContent.substring(0, 100) + '...');
        this.setState({
          htmlContent: trimmedContent, // Use trimmed content
          isLoading: false,
          error: ''
        });
        console.log('✅ Flow executed successfully with', authMethod);
        
        // Force a re-render to ensure the content shows up
        setTimeout(() => {
          console.log('🔄 Forcing re-render after state update');
          this.renderContent();
        }, 100);
      }
    } else {
      const errorBody = await response.text().catch(() => 'Could not read error details');
      const errorMessage = `HTTP ${response.status}: ${response.statusText}`;
      
      throw new Error(`${errorMessage} (${authMethod}): ${errorBody}`);
    }
  }

  private showPnpAuthenticationFailure(flowUrl: string, pnpError: string, fetchError: string): void {
    this.setState({
      error: 'PnP authentication failed',
      htmlContent: `
        <div style="padding: 20px; border: 1px solid #dc3545; background-color: #f8d7da;">
          <h3>PnP Authentication Failed</h3>
          <p>Unable to call the Power Automate flow using the PnP recommended approach.</p>
          <p><strong>Flow URL:</strong> ${escape(flowUrl)}</p>
          
          <details style="margin: 10px 0;">
            <summary>Error Details</summary>
            <div style="background: #f8f9fa; padding: 10px; border-radius: 4px; margin: 5px 0;">
              <h6>PnP AadTokenProvider Error:</h6>
              <pre style="overflow-x: auto; max-height: 100px;">${escape(pnpError)}</pre>
            </div>
            <div style="background: #f8f9fa; padding: 10px; border-radius: 4px; margin: 5px 0;">
              <h6>Direct Fetch Error:</h6>
              <pre style="overflow-x: auto; max-height: 100px;">${escape(fetchError)}</pre>
            </div>
          </details>
          
          <h4>🔧 Solutions (PnP Recommended):</h4>
          
          <div style="background: #e7f3ff; padding: 15px; border-radius: 4px; margin: 10px 0; border-left: 4px solid #007acc;">
            <h5>✅ Step 1: Verify API Permissions</h5>
            <p>Ensure these permissions are approved in SharePoint Admin Center:</p>
            <pre style="background: #f8f9fa; padding: 10px; border-radius: 4px; overflow-x: auto;">
"webApiPermissionRequests": [
  {
    "resource": "Microsoft Flow Service",
    "scope": "User"
  }
]</pre>
            <p><strong>How to approve:</strong></p>
            <ol>
              <li>Go to SharePoint Admin Center</li>
              <li>Navigate to Advanced > API Access</li>
              <li>Approve all pending requests</li>
              <li>Refresh this page</li>
            </ol>
          </div>
          
          <div style="background: #f8f9fa; padding: 15px; border-radius: 4px; margin: 10px 0;">
            <h5>Option 2: Use Public HTTP Trigger (Alternative)</h5>
            <ol>
              <li>In Power Automate, create a new flow</li>
              <li>Use "When an HTTP request is received" trigger</li>
              <li>Set "Who can trigger this flow" to "Anyone"</li>
              <li>This bypasses authentication requirements</li>
              <li>Add proper JSON schema for request validation</li>
              <li>End with a "Response" action that returns HTML</li>
            </ol>
          </div>
          
          <div style="background: #f8f9fa; padding: 15px; border-radius: 4px; margin: 10px 0;">
            <h5>Option 3: Verify Dependencies</h5>
            <p>Ensure jwt-decode is properly installed:</p>
            <pre style="background: #f8f9fa; padding: 5px; border-radius: 4px;">npm install jwt-decode @types/jwt-decode</pre>
          </div>
        </div>`,
      isLoading: false
    });
  }
}
