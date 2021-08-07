import { ComprehendClient, DetectKeyPhrasesCommand, DetectKeyPhrasesCommandInput, KeyPhrase } from "@aws-sdk/client-comprehend";
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

    const detectKeyPhrasesCommand: DetectKeyPhrasesCommand = new DetectKeyPhrasesCommand(params);

    returnData = await client.send(detectKeyPhrasesCommand);

    returnData.KeyPhrases.reverse().forEach((keyPhrase: KeyPhrase) => {

      if (keyPhrase.Score > 0.999) {

        // last highlight
        var b = "</mark>";
        var position = keyPhrase.EndOffset

        this.emailBody = [this.emailBody.slice(0, position), b, this.emailBody.slice(position)].join('');

        // first highlight
        var b = "<mark>";
        var position = keyPhrase.BeginOffset;
        this.emailBody = [this.emailBody.slice(0, position), b, this.emailBody.slice(position)].join('');
      }
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
