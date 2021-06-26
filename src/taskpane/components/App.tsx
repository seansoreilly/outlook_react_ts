import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import * as Functions from "./Functions";
import * as Aws from "./Aws";
import * as Commands from "../../commands/commands";

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

    var configAws = { invokeUrl: 'https://comprehend.us-east-2.amazonaws.com/' }
    var apigClient = Aws.apigClientFactory.newClient(configAws);
    var _result;
    var pathTemplate = '/users/{userID}/profile'
    var method = 'GET';    

    var params = {
      // This is where any modeled request parameters should be added.
      // The key is the parameter name, as it is defined in the API in API Gateway.
      param0: '',
      param1: ''
    };

    var body = {
      // This is where you define the body of the request,
      "LanguageCode": "en",
      "Text": "{\n    \"Text\": \"TO THE CHIEF EXECUTIVE OFFICER Sent to CEOs and CEO EAs  Dear Colleagues, MAV State Council Meeting – 'Save the Date' – Friday 21 May 2020 The next MAV State Council Meeting will be held from 9.30am to 2:30pm on Friday 21 May 2021 at the Melbourne Town Hall, Corner Swanston and Collins Streets Melbourne.  The online links will be available to submit motions and register to attend in mid-March and in accordance with the MAV Rules, motions are to be submitted no later than midnight on Friday 23 April 2021.\n}"
    };

    var additionalParams = {
      // If there are any unmodeled query parameters or headers that must be
      //   sent with the request, add them here.
      headers: {
        param0: '',
        param1: ''
      },
      queryParams: {
        param0: '',
        param1: ''
      }
    };

    apigClient.invokeApi(params, method, pathTemplate, additionalParams, body)
      .then(function (_result) {
        // Add success callback code here.
        console.log("success");
      }).catch(function (_result) {
        // Add error callback code here.
        console.log("error");
      });

    console.log(_result);

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
