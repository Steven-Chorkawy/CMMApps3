export const emailRegex = new RegExp(/\S+@\S+\.\S+/);
export const emailValidator = (value: any): string => (value === undefined || emailRegex.test(value)) ? "" : "Please enter a valid email.";