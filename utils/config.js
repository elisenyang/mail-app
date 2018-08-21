module.exports = {
    creds: {
        identityMetadata: 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
        clientID: process.env.clientID, //MOVE THIS LATER
        responseType: 'code',
        responseMode: 'form_post',
        redirectUrl: 'http://localhost:3000/token',
        allowHttpForRedirectUrl: true,
        clientSecret: process.env.clientSecret, //MOVE THIS LATER
        validateIssuer: false,
        scope: ['User.Read', 'Mail.Read', 'Mail.ReadWrite', 'User.ReadWrite', 'profile']
    }
};