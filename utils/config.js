module.exports = {
    creds: {
        identityMetadata: 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
        clientID: "25ac6d4b-c5ca-476a-867f-bab0b52e3405",
        responseType: 'code',
        responseMode: 'form_post',
        redirectUrl: 'http://localhost:3000/token',
        allowHttpForRedirectUrl: true,
        clientSecret: "vqjmrAV41@%rvJKSMW449:~", //MOVE THIS LATER
        validateIssuer: false,
        scope: ['User.Read', 'Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'User.ReadWrite', 'profile']
    }
};