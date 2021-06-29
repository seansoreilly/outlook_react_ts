// https://docs.aws.amazon.com/AWSJavaScriptSDK/v3/latest/clients/client-comprehend/index.html

import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import * as Functions from "./Functions";
// import * as Aws from "./Aws";
import * as Commands from "../../commands/commands";

import { ComprehendClient, BatchDetectDominantLanguageCommand, DetectSentimentCommand, DetectSentimentRequest } from "@aws-sdk/client-comprehend"; // ES Modules import

// a client can be shared by different commands.
const client = new ComprehendClient({ region: "REGION" });
const params = {
  /** input parameters */
};

const command = new BatchDetectDominantLanguageCommand(params);

const response_await = async () => {
  const response = await client.send(command);
};

// image references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  click = async () => {

    /**
     * Insert your Outlook code here
     */

    // var configAws = { invokeUrl: 'https://comprehend.us-east-2.amazonaws.com/' }
    // var apigClient = Aws.apigClientFactory.newClient(configAws);
    // var _result = '{}';
    // // var pathTemplate = '/users/{userID}/profile'
    // var pathTemplate = ''
    // var method = 'POST';   

    // var params = {
    //   // This is where any modeled request parameters should be added.
    //   // The key is the parameter name, as it is defined in the API in API Gateway.
    //   param0: '',
    //   param1: ''
    // };

    // var body = {
    //   // This is where you define the body of the request,
    //   "LanguageCode": "en",
    //   "Text": "TO THE CHIEF EXECUTIVE OFFICER"
    // };

    // var additionalParams = {
    //   // If there are any unmodeled query parameters or headers that must be
    //   //   sent with the request, add them here.
    //   'headers': {
    //     // 'Service': 'Comprehend',
    //     'X-amz-target': 'Comprehend_20171127.DetectSentiment',
    //     'Content-Type': 'application/x-amz-json-1.1'
    //     // 'X-Amz-Content-Sha256': 'beaead3198f7da1e70d03ab969765e0821b24fc913697e929e726aeaebf0eba3'
    //     // 'X-Amz-Date': '20210629T111743Z',
    //     // 'Authorization': 'AWS4-HMAC-SHA256 Credential=AKIASGQWQXVAHSKVCJVD/20210629/us-east-2/comprehend/aws4_request, SignedHeaders=content-type;host;x-amz-content-sha256;x-amz-date;x-amz-target, Signature=2718d416061f641d4bbdd9735aa8ef322f4089f59afde1fcc3a89892193b4269'
    //   }
    //   // queryParams: {
    //   //   param0: '',
    //   //   param1: ''
    //   // }
    // };

    // // apigClient.invokeApi(params, pathTemplate, method, additionalParams, body)
    // apigClient.invokeApi('', '', method, additionalParams, body)
    //   .then(function (_result) {
    //     // Add success callback code here.
    //     console.log("success");
    //     console.log(_result);
    //   }).catch(function (_result) {
    //     // Add error callback code here.
    //     console.log("error");
    //   });

    // // console.log(JSON.parse(_result));

    var getSalutation: string = Functions.salutation(Office.context.mailbox.item.to);
    console.log(getSalutation);
    Commands.putNotificationMessage(getSalutation);

  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList>
      </div>
    );
  }
}
