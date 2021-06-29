var config = require("../../../config.js");

export var apigClientFactory = require('aws-api-gateway-client').default;

export var apigClient = apigClientFactory.newClient({
    invokeUrl: 'https://comprehend.us-east-2.amazonaws.com/',        // REQUIRED
    region: config.AWS_REGION,                                           // REQUIRED: The region where the API is deployed.
    accessKey: config.AWS_ACCESS_KEY,                                       // REQUIRED
    secretKey: config.AWS_SECRET_KEY,                                       // REQUIRED
    // sessionToken: 'SESSION_TOKEN',                                 // OPTIONAL: If you are using temporary credentials you must include the session token.
    systemClockOffset: 0,                                          // OPTIONAL: An offset value in milliseconds to apply to signing time
    retries: 4,                                                    // OPTIONAL: Number of times to retry before failing. Uses axios-retry plugin.
    // retryCondition: (err) => {                                     // OPTIONAL: Callback to further control if request should be retried.
    //     return err.response && err.response.status === 500;          //           Uses axios-retry plugin.
    // },

    // retryDelay: 100 || 'exponential' || (5, error) => {   // OPTIONAL: Define delay (in ms) as a number, a callback, or
    // return retryCount * 100                                      //           'exponential' to use the in-built exponential backoff
    // },                                                             //           function. Uses axios-retry plugin. Default is no delay.
    shouldResetTimeout: false                                      // OPTIONAL: Defines if the timeout should be reset between retries. Unless
    //           `shouldResetTimeout` is set to `true`, the request timeout is
    //           interpreted as a global value, so it is not used for each retry,
    //           but for the whole request lifecycle.
});


// var apigClient = apigClientFactory.newClient({
//     invokeUrl: 'https://comprehend.us-east-2.amazonaws.com/', // REQUIRED
//     apiKey: 'API_KEY', // REQUIRED
//     region: 'us-east-2' // REQUIRED
// });
