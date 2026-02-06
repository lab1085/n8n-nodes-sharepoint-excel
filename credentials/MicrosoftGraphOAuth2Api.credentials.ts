import type { ICredentialTestRequest, ICredentialType, INodeProperties, Icon } from 'n8n-workflow';

export class MicrosoftGraphOAuth2Api implements ICredentialType {
	name = 'microsoftGraphOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	displayName = 'Microsoft Graph OAuth2 API';

	documentationUrl = 'https://learn.microsoft.com/en-us/graph/auth-v2-user';

	icon: Icon = 'file:icons/Microsoft.svg';

	properties: INodeProperties[] = [
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default: 'openid offline_access Sites.Read.All Files.ReadWrite.All',
		},
	];

	test: ICredentialTestRequest = {
		request: {
			baseURL: 'https://graph.microsoft.com/v1.0',
			url: '/sites?search=*&$top=1',
		},
	};
}
