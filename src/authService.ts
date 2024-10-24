import { AuthenticationResult, createNestablePublicClientApplication, IPublicClientApplication } from '@azure/msal-browser';

export interface IAuthService {
	getIdToken(scopes?: string[]): Promise<string | null>;
	getUserIdentity(scopes?: string[]): Promise<AuthenticationResult | null>;
}

export class AuthService implements IAuthService {
	private constructor(private readonly pca: IPublicClientApplication) {}

	public static async createAsync(appId?: string): Promise<IAuthService> {
		const applicationId = appId ?? process.env.ADD_IN_APP_ID ?? '';
		const pca = await createNestablePublicClientApplication({
			auth: {
				clientId: applicationId,
				authority: 'https://login.microsoftonline.com/common',
			},
		});

		return new AuthService(pca);
	}

	private isNestedAppAuthSupported(): boolean {
		const naaSupported: boolean = Office.context.requirements.isSetSupported('NestedAppAuth', '1.1');
		return naaSupported;
	}

	public async getIdToken(scopes: string[] = ['openid', 'profile']): Promise<string | null> {
		const userAccount = await this.getUserIdentity(scopes);
		return userAccount?.idToken ?? null;
	}

	public async getUserIdentity(scopes: string[]): Promise<AuthenticationResult | null> {
		if (!this.isNestedAppAuthSupported()) {
			console.warn('NestedAppAuth is not supported.');
			return null;
		}

		const tokenRequest = {
			scopes,
		};

		try {
			// next line is the last line that will be executed
			const authResult: AuthenticationResult = await this.pca.acquireTokenSilent(tokenRequest);

			// this will never be executed
			console.log('Silent token acquisition successful.', authResult);
			return authResult;
		} catch (silentError) {
			// this will never be executed
			console.error('Silent token acquisition failed.', silentError);
		}

		return null;
	}
}
