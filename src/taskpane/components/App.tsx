import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
// import * as Functions from "./Functions";
// import * as Commands from "../../commands/commands";
require("../../../config.js");
// declare var emailBody: string;

import { ComprehendClient, DetectKeyPhrasesCommand, DetectKeyPhrasesCommandInput } from "@aws-sdk/client-comprehend";

/* global Outlook, Office, OfficeExtension */

// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
// import { ResolvePlugin } from "webpack";

// global variables
let returnData: any;

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

    const creds = {
      accessKeyId: process.env.accessKeyId,
      secretAccessKey: process.env.secretAccessKey
    };

    let emailBody: string = await getBody().then(function (result) {
      return result;
    });

    const client = new ComprehendClient({ region: process.env.region, credentials: creds });

    const params: DetectKeyPhrasesCommandInput = {
      LanguageCode: "en",
      Text: emailBody
    };

    const command = new DetectKeyPhrasesCommand(params);

    client.send(command).then(
      (data) => {
        returnData = data;
      },
      (error) => {
        console.log(error)
      }
    );

    // console.log(emailBody);

    returnData.KeyPhrases.reverse().forEach(KeyPhrase => {
      console.log(KeyPhrase);
      // last bold
      var b = "</mark>";
      var position = KeyPhrase.EndOffset;
      emailBody = [emailBody.slice(0, position), b, emailBody.slice(position)].join('');

      // first bold
      var b = "<mark>";
      var position = KeyPhrase.BeginOffset;
      emailBody = [emailBody.slice(0, position), b, emailBody.slice(position)].join('');


    });

    console.log(emailBody);

    // put text back into email

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

function getBody(): Promise<string> {

  // return new OfficeExtension.Promise(function (resolve) {
  return new Office.Promise(function (resolve, reject) {

    try {
      Office.context.mailbox.item.body.getAsync(
        'text',
        function (asyncResult) {
          resolve(asyncResult.value)
        }
      )
    }

    catch (error) {
      console.log(error.toString());
      reject(error.toString());
    }

    finally {
    }
  });
}
