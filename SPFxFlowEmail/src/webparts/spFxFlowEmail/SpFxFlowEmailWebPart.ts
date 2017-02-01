import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './SpFxFlowEmail.module.scss';
import * as strings from 'spFxFlowEmailStrings';
import { ISpFxFlowEmailWebPartProps } from './ISpFxFlowEmailWebPartProps';

import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
interface IResult {
  status: string;
}

export default class SpFxFlowEmailWebPart extends BaseClientSideWebPart<ISpFxFlowEmailWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.row}">
        <div class="${styles.column}">
          <span class="${styles.title}">
            SPFx Flow Email web part
          </span>
          <p class="${styles.subtitle}">
            Compliments of <a href="https://mvp.microsoft.com/en-us/PublicProfile/36831?fullName=Todd%20S%20Baginski" target="_blank">Todd Baginski</a> and <a href="http://www.canviz.com" target="_blank">Canviz</a>
          </p>
          <p class="${styles.description}">
            Learn more about this web part on my blog <a href="http://toddbaginski.com/blog/how-to-run-a-microsoft-flow-from-a-sharepoint-framework-spfx-web-part" target="_blank">http://toddbaginski.com/blog/how-to-run-a-microsoft-flow-from-a-sharepoint-framework-spfx-web-part</a>
          </p>
          <p class="${styles.description}">
            Open the property pane to configure the web part to send email through a Microsoft Flow.
          </p>
          <button class="ms-Button ${styles.button}">
            Send Email
          </button>
        </div>
      </div>`;

    const buttons = this.domElement.querySelectorAll('button');

    buttons[0].addEventListener("click", (evt: Event): void => {
      this.sendEmail();
      evt.preventDefault();
    });
  }

  private validateEmail(email) {
    var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
  }

  private sendEmailViaOffice365Outlook(emailaddress: string, emailSubject: string, emailBody: string): Promise<HttpClientResponse> {

    const postURL = this.properties.flowURL;

    const body: string = JSON.stringify({
      'emailaddress': emailaddress,
      'emailSubject': emailSubject,
      'emailBody': emailBody,
    });

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');

    const httpClientOptions: IHttpClientOptions = {
      body: body,
      headers: requestHeaders
    };

    console.log("Sending Email");

    return this.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      httpClientOptions)
      .then((response: Response): Promise<HttpClientResponse> => {
        console.log("Email sent.");
        return response.json();
      });
  }

  private sendEmail() {
    if (!this.properties.flowURL) {
      this.context.statusRenderer.renderError(this.domElement, "Flow URL is missing.  Open the property pane and specify a Flow URL.");
      return;
    }
    if (!this.validateEmail(this.properties.emailAddress)) {
      this.context.statusRenderer.renderError(this.domElement, "Email address is not valid.  Open the property pane and specify a valid email address.");
      return;
    }
    else {
      this.sendEmailViaOffice365Outlook(this.properties.emailAddress,
        this.properties.emailSubject,
        this.properties.emailBody);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            {
              groupName: strings.EmailGroupName,
              groupFields: [
                PropertyPaneTextField('flowURL', {
                  label: strings.FlowURLLabel
                }),
                PropertyPaneTextField('emailAddress', {
                  label: strings.EmailAddressFieldLabel
                }),
                PropertyPaneTextField('emailSubject', {
                  label: strings.EmailSubjectFieldLabel
                }),
                PropertyPaneTextField('emailBody', {
                  label: strings.EmailBodyFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
