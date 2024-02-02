import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAltayerpopupProps {
  description: string;
  userDisplayName: string;
  webpartContext:WebPartContext;
  listName:string;
  responseListName:string;
  consentTerms:string;
  neverShowText:string;
}
