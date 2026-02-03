import type { ICredentialType, INodeProperties } from 'n8n-workflow';

export class MicrosoftGraphOAuth2Api implements ICredentialType {
	name = 'microsoftGraphOAuth2Api';

	extends = ['oAuth2Api'];

	displayName = 'Microsoft Graph OAuth2 API';

	documentationUrl = 'https://docs.microsoft.com/en-us/graph/auth-v2-user';

	icon = 'file:icons/Microsoft.svg' as const;

	properties: INodeProperties[] = [
		{
			displayName: 'Grant Type',
			name: 'grantType',
			type: 'hidden',
			default: 'authorizationCode',
		},
		{
			displayName: 'Authorization URL',
			name: 'authUrl',
			type: 'hidden',
			default: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
		},
		{
			displayName: 'Access Token URL',
			name: 'accessTokenUrl',
			type: 'hidden',
			default: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
		},
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default: 'openid offline_access Sites.ReadWrite.All Files.ReadWrite.All',
		},
		{
			displayName: 'Auth URI Query Parameters',
			name: 'authQueryParameters',
			type: 'hidden',
			default: 'response_mode=query',
		},
		{
			displayName: 'Authentication',
			name: 'authentication',
			type: 'hidden',
			default: 'body',
		},
	];
}
