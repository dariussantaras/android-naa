import { AuthService } from "../authService";

export async function handleOnMessageComposeAsync(): Promise<boolean> {
	const authService = await AuthService.createAsync(process.env.ADD_IN_APP_ID);

	const token = await authService.getIdToken();

	Office.context.mailbox.item?.body.setSignatureAsync(`token: ${token}`, {}, (result) => {
		result.status === Office.AsyncResultStatus.Succeeded ? console.log("Signature set") : console.error(result.error);
	});

	return true;
}
