import * as React from "react";
import { ActionButton, ButtonType } from "office-ui-fabric-react";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import DetectKeyPhrases from "./DetectKeyPhrases/DetectKeyPhrases";

/* global Outlook, Office, OfficeExtension */

// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import ReactDOM = require("react-dom");

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

    const DKP = new DetectKeyPhrases();
    await DKP.getKeyPhrases();
    let emailBody: string = DKP.emailBody;

    const emailBodyHTML = <div className="Container" dangerouslySetInnerHTML={{ __html: emailBody }}></div>;

    ReactDOM.render(emailBodyHTML, document.getElementById('displayResult'));

  };

  render() {

    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div>
        <div className="ms-welcome">

          <ActionButton
            className="ms-welcome__action"
            // buttonType={ButtonType.hero}
            buttonType={ButtonType.command}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </ActionButton >

        </div>
        <div id="displayResult" className="ms-welcome__html" >
        </div>
      </div>
    );
  }
}

// function getBody(): Promise<string> {

//   return new Office.Promise(function (resolve, reject) {

//     try {
//       Office.context.mailbox.item.body.getAsync(
//         'text',
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
