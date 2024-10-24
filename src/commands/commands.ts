// import { AuthService, IAuthService } from '../authService';
import { handleOnMessageComposeAsync } from './onMessageCompose';

// let authService: IAuthService;

Office.onReady(async () => {
	// authService = await AuthService.createAsync(process.env.ADD_IN_APP_ID);
});

export function onMessageComposeHandler(event: Office.AddinCommands.Event): Promise<void> {
	return handleEventAsync(event, () => handleOnMessageComposeAsync());
}

Office.actions.associate('onMessageComposeHandler', onMessageComposeHandler); 

async function handleEventAsync<T extends Office.AddinCommands.Event>(event: T, handler: (event: T) => Promise<boolean>): Promise<void> {
	try {
		await handler(event);
	} catch (error) {
        console.error(error);
	} finally {
		event.completed();
	}
}
