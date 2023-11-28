# IMAPOAuthSample

This is a sample console application written in .Net Core that demonstrates how to obtain an OAuth token for logging on to a mailbox using IMAP.  Note that IMAP is a public protocol and as such it is up to the developer to correctly implement it in their code. The example here is basic, and only intended to show how OAuth fits in to the log-in process. You do not have to use MSAL to obtain the token, but it is a very simple way to do so.

You must register the application in Azure AD as per [this guide](https://docs.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#get-an-access-token "Authenticate an IMAP application using OAuth").  By default, the application will specify http://localhost as the redirect URL, but you can use your own and then pass that in as a parameter in the delegated auth flow.

Once the application is registered, the application can be run from a command prompt (or PowerShell console).  The syntax is:

## Delegated authentication:
```
IMAPOAuthSample TenantId ApplicationId (RedirectURL)
```

If redirect URL is not specified, it will default to http://localhost.
Delegated authentication requires the user to log in, which they will be prompted to do via a web browser.

## Authenticate as application:
```
IMAPOAuthSample TenantId ApplicationId SecretKey Mailbox
```

Application authentication does not require user input for log-in, but you must specify the secret key and mailbox to be accessed.

## Example

Once a valid token is retrieved (via either flow), the application will use it to log on to the mailbox and retrieve the number of unread messages in the Inbox.  The IMAP conversation will be shown in the console.

A successful test (delegate auth) looks like this:

![IMAPOAuthSample Successful Test Screenshot](https://github.com/David-Barrett-MS/IMAPOAuthSample/blob/master/IMAPOAuthSample.png?raw=true "IMAPOAuthSample Successful Test Screenshot")
