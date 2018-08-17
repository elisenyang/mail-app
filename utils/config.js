module.exports = {
    creds: {
        identityMetadata: 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
        clientId: '25ac6d4b-c5ca-476a-867f-bab0b52e3405', //MOVE THIS LATER
        responseType: 'code',
        responseMode: 'form_post',
        redirectURL: 'http://localhost:3000/token',
        passReqToCallback: true,
        allowHttpForRedirectUrl: true,
        clientSecret: 'xpBR745#}+ydrzwOGCXF77{', //MOVE THIS LATER
        validateIssuer: false,
        scope: ['User.Read', 'Mail.Read', 'Mail.ReadWrite',  'profile']
    }
};