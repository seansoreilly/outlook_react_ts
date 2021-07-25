import { ComprehendClient, DetectKeyPhrasesCommand, DetectKeyPhrasesCommandInput } from "@aws-sdk/client-comprehend";
require("../../../../config.js");

export default class DetectKeyPhrases {

  emailBody: string = "";

  public async getKeyPhrases() {

    let returnData: any;

    const creds = {
      accessKeyId: process.env.accessKeyId,
      secretAccessKey: process.env.secretAccessKey
    };

    this.emailBody = await getBody().then(function (result) {
      return result;
    });

    const client = new ComprehendClient({ region: process.env.region, credentials: creds });

    const params: DetectKeyPhrasesCommandInput = {
      LanguageCode: "en",
      Text: this.emailBody
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

    returnData.KeyPhrases.reverse().forEach(KeyPhrase => {
      console.log(KeyPhrase);
      // last bold
      var b = "</mark>";
      var position = KeyPhrase.EndOffset;
      this.emailBody = [this.emailBody.slice(0, position), b, this.emailBody.slice(position)].join('');

      // first bold
      // var b = "<mark>";
      // var position = KeyPhrase.BeginOffset;
      // emailBody = [emailBody.slice(0, position), b, emailBody.slice(position)].join('');

    });

    return this.emailBody;

  };

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

