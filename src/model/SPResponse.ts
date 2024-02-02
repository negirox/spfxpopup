export interface SPResponse {
    "@odata.context"?: string
    value: ResponseValue[]
  }
  
  export interface ResponseValue {
    "@odata.type": string
    "@odata.id": string
    "@odata.etag": string
    "@odata.editLink": string
    Id: number
    Title: string
    Description: string
    ExpiryDate: string
  }
  