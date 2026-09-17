export type Submission = {
  id: number
  submissionId: string
  nsc: string
  functionName: string
  category: string
  subcategory: string
  contractingParty: string
  contractValidity: string
  reviewer: string
  fulfiller: string
  submitter: string
  submittedDate: string
  status: string
}

export type FilterValues = {
  nsc: string
  functionName: string
  category: string
  subcategory: string
  contractingParty: string
  status: string
}
