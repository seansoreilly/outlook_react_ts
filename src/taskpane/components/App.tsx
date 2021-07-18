import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import * as Functions from "./Functions";
import * as Commands from "../../commands/commands";
require("../../../config.js");

import { ComprehendClient, DetectKeyPhrasesCommand } from "@aws-sdk/client-comprehend";

// images references in the manifest
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

    console.log(process.env.accessKeyId);

    const creds = {
      accessKeyId: process.env.accessKeyId,
      secretAccessKey: process.env.secretAccessKey
    };

    var emailBody;

    // const emailBody = "Further to our chat on Wednesday, attached is a draft alternate motion for the above application that is to be considered on Monday night.   As discussed, I have added a condition requiring a Waste Management Plan (condition 3) that among other matters requires the development to utilise a shared bin service, which will reduce the number of bins required by a considerable number.  Please let me know if you are OK with the alternate as drafted, or if you would like any changes made.";
    // Office.context.mailbox.item.body.getAsync('text', function (async) {const emailBody = async.value)});
    Office.context.mailbox.item.body.getAsync(
      'text',
      function (async) { emailBody = async.value }
    );

    const client = new ComprehendClient({ region: process.env.region, credentials: creds });

    const params = {
      "LanguageCode": "en",
      "Text": emailBody
    };

    const command = new DetectKeyPhrasesCommand(params);

    client.send(command).then(
      (data) => {
        console.log(data);
        data.KeyPhrases.forEach(element =>

          console.log(element)

        );
      },
      (error) => {
        console.log(error)
      }
    );








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
