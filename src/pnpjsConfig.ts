/*eslint-disable*/ 
/*eslint no-empty-function: ["warn", { "allow": ["functions","arrowFunctions"] }]*/
/*eslint-env es6*/
// tslint:disable-next-line:no-empty
import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp and pnp logging system
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";
let _sp: SPFI = null;
export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context !== null) {
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
}
