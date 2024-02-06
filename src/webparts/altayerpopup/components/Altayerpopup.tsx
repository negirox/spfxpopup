import * as React from 'react';
import styles from './Altayerpopup.module.scss';
import { IAltayerpopupProps } from './IAltayerpopupProps';
import {
  SPHttpClient
} from '@microsoft/sp-http';
import { FocusTrapZone, Layer, Popup, PrimaryButton, mergeStyleSets } from 'office-ui-fabric-react';
import { SPResponse } from '../../../model/SPResponse';
import { AlertModel } from '../../../model/IAlertModel';
import { getSP } from '../../../pnpjsConfig';
import { SPFI } from '@pnp/sp';
//import { escape } from '@microsoft/sp-lodash-subset';
const popupStyles = mergeStyleSets({
  root: {
    background: 'rgba(0, 0, 0, 0.2)',
    bottom: '0',
    left: '0',
    position: 'fixed',
    right: '0',
    top: '0',
  },
  content: {
    background: 'white',
    left: '50%',
    maxWidth: '70%',
    padding: '1.5rem',
    position: 'absolute',
    top: '50%',
    transform: 'translate(-50%, -50%)',
    borderTop: '4px solid rgb(0, 85, 150)',
    maxHeight: '40%',
    overflow: 'auto',
    scrollbar: "thin",
    fontSize: '12px'
  },
});
export default class Altayerpopup extends React.Component<IAltayerpopupProps, {
  isPopupVisible: boolean,
  records: Array<AlertModel>,
  acceptConsent: boolean,
  userResponseItemId: number,
  isAlwaysView: boolean
}> {
  private _spContext: SPFI;
  private _loggedInUserId: number;
  constructor(props: IAltayerpopupProps) {
    super(props);
    this.state = {
      isPopupVisible: false,
      records: new Array<AlertModel>(),
      acceptConsent: false,
      userResponseItemId: 0,
      isAlwaysView: false
    }
    this.handleAgreement = this.handleAgreement.bind(this);
    this.AcceptConsent = this.AcceptConsent.bind(this);
    this.AcceptDontShow = this.AcceptDontShow.bind(this);
    this._spContext = getSP(this.props.webpartContext);
  }
  async componentDidMount(): Promise<void> {
    const items: SPResponse = await this._getListData();
    if (items.value.length > 0) {
      const ConfigUrl = `${this.props.webpartContext.pageContext.web.absoluteUrl}/_api/web/currentuser?$select=Title,Email,Id`;
      const response = await this.props.webpartContext.spHttpClient.get(ConfigUrl, SPHttpClient.configurations.v1);
      const responseValue: any = await response.json();
      this._loggedInUserId = responseValue.Id;
      await this.ConvertToModel(items);
    }
    else {
      //if no records return disable the popup
      this.setState({
        records: [],
        isPopupVisible: false
      });
    }
    return Promise.resolve();
  }
  private async ConvertToModel(items: SPResponse): Promise<void> {
    const records = items.value;
    const alertModelRecords = new Array<AlertModel>();
    let ConsentItemId = 0;
    let isRecorded = false;
    records.forEach(x => {
      const model = new AlertModel();
      model.Id = x.Id;
      model.Description = x.Description;
      model.ExpiryDate = x.ExpiryDate;
      model.Title = x.Title;
      alertModelRecords.push(model);
      ConsentItemId = x.Id;
    });
    if (ConsentItemId > 0) {
      const response = await this.IsResponseRecorded(ConsentItemId);
      if (response.length > 0) {
        if (response[0].Id > 0) {
          isRecorded = response[0].NeverShow;
          this.setState({
            userResponseItemId: response[0].Id,
            isAlwaysView: response[0].NeverShow
          })
        }
      }
    }

    this.setState({
      records: alertModelRecords,
      isPopupVisible: !isRecorded
    });
  }
  private async IsResponseRecorded(ConsentItemId: number): Promise<any[]> {
    const response = await this._spContext.web.lists.getByTitle(this.props.responseListName).items.filter(
      `UserEmail eq '${this.props.webpartContext.pageContext.user.email}' and ConsentId eq '${ConsentItemId}'`
    ).top(1).select('Id,NeverShow')();
    return response;
  }
  private getFormattedDate(date: Date): string {
    const year = date.getFullYear();

    let month = (1 + date.getMonth()).toString();
    month = month.length > 1 ? month : '0' + month;

    let day = date.getDate().toString();
    day = day.length > 1 ? day : '0' + day;
    const amOrPm = (date.getHours() < 12) ? "AM" : "PM";
    const hour = (date.getHours() < 12) ? '0' + date.getHours() : date.getHours() - 12;
    return month + '/' + day + '/' + year + ' ' + hour + ':' + date.getMinutes() + ' ' + amOrPm;
  }
  private async _getListData(): Promise<SPResponse> {
    const records = `&$top=4999`;
    const selectedColumns = `Title,Id,Description,ExpiryDate`;
    const isFutureEvent = 'ge';
    const filterCondition = `ExpiryDate ${isFutureEvent} '${new Date().toISOString()}'`;
    const orderByColumn = `ExpiryDate desc`;
    const ConfigUrl = `${this.props.webpartContext.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/Items?$filter=${filterCondition}&$select=${selectedColumns}${records}&$orderby=${orderByColumn}`;
    const response = await this.props.webpartContext.spHttpClient.get(ConfigUrl, SPHttpClient.configurations.v1);
    const responseValue: SPResponse = await response.json();
    return responseValue;
  }
  private async handleAgreement(): Promise<void> {
    if (this.state.records.length > 0) {
      const postObj = {
        UserEmail: this.props.webpartContext.pageContext.user.email,
        ConsentId: this.state.records[0].Id?.toString(),
        UserNameId: this._loggedInUserId,
        ResponseTime: this.getFormattedDate(new Date()),
        NeverShow: this.state.isAlwaysView
      }
      if (this.state.userResponseItemId === 0) {
        const result = await this._spContext.web.lists.getByTitle(this.props.responseListName).items.add(postObj);
        if (result.item) {
          console.log(result);
        }
      }
      else {
        const result = await this._spContext.web.lists.getByTitle(this.props.responseListName).items
          .getById(this.state.userResponseItemId).update(postObj);
        if (result.item) {
          console.log(result);
        }
      }

      this.setState({
        isPopupVisible: false
      })
    }
    else {
      console.log('no data found');
    }
  }
  private AcceptConsent(): void {
    this.setState({
      acceptConsent: !this.state.acceptConsent
    });
  }
  private AcceptDontShow(): void {
    this.setState({
      isAlwaysView: !this.state.isAlwaysView
    });
  }
  public render(): React.ReactElement<IAltayerpopupProps> {

    return (
      <section className={styles.altayerpopup}>
        <div>
          <Layer>
            {this.state.isPopupVisible && <Popup
              className={popupStyles.root}
              role="dialog"
              aria-modal="true"
            >
              <FocusTrapZone>
                {this.state.records.length > 0 && <div role="document" className={popupStyles.content}>
                  <h2>
                    {
                      this.state.records.length > 0 && this.state.records[0].Title
                    }
                  </h2>

                  <p className='flex'
                    dangerouslySetInnerHTML={{ __html: this.state.records.length > 0 && this.state.records[0].Description }} />
                  <div className='flex'>
                    <div>
                      <label style={{ fontWeight: 700 }}>
                        <input type="checkbox" onChange={this.AcceptConsent} />
                        {(this.props.consentTerms)}
                      </label>
                    </div>
                    <div style={{ marginTop: '.4rem' }}>
                      <label style={{ fontWeight: 700 }}>
                        <input type="checkbox" onChange={this.AcceptDontShow} />
                        {(this.props.neverShowText)}
                      </label>
                    </div>
                  </div>
                  <div style={{ marginTop: '1rem', textAlign:'center' }}>
                    <PrimaryButton onClick={this.handleAgreement}
                      disabled={!this.state.acceptConsent}>I Agree</PrimaryButton>
                  </div>
                </div>
                }
              </FocusTrapZone>
            </Popup>
            }
          </Layer>
        </div>
      </section>
    );
  }
}
