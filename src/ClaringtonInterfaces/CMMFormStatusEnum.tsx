export enum CMMFormStatus {
    NewForm = 1,    // New form/ Cleared Form. 
    Processing = 2, // Form has been submitted waiting for result.
    Success = 3,    // Form successfully submitted.
    Error = 4       // Form failed to submit.
}