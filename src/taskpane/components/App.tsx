import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import DetectKeyPhrases from "./DetectKeyPhrases/DetectKeyPhrases";
// import * as Functions from "./Functions";
// import * as Commands from "../../commands/commands";
// require("../../../config.js");
// declare var emailBody: string;

import { ComprehendClient, DetectKeyPhrasesCommand, DetectKeyPhrasesCommandInput } from "@aws-sdk/client-comprehend";

/* global Outlook, Office, OfficeExtension */

// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
// import { ResolvePlugin } from "webpack";

// global variables
// let returnData: any;

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

    //figure out how to call the function from here
    const DKP = new DetectKeyPhrases();
    let emailBody2: any = DKP.getKeyPhrases();
    //working OK, but returning promise rather than value




    console.log(emailBody2);

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

// NB: not used at the moment
// function putBody(sBody: string): Promise<any> {

//   return new Office.Promise(function (resolve, reject) {

//     try {
//       Office.context.mailbox.item.body.setAsync(
//         sBody,
//         { coercionType: Office.CoercionType.Html },
//         function (asyncResult) {
//           resolve(asyncResult.value)
//         }
//       )
//     }

//     catch (error) {
//       console.log(error.toString());
//       reject(error.toString());
//     }

//     finally {
//     }

//   });
// }