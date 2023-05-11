import * as React from 'react';
import styles from './Pnpcontrols.module.scss'; ;
import { IPnpcontrolsProps } from './IPnpcontrolsProps';
import { IPnpcontrolsState} from './IPnpcontrolsState';
import { escape } from '@microsoft/sp-lodash-subset';
import {PeoplePicker} from '@pnp/spfx-controls-react/lib/PeoplePicker';
import {IPickerTerms, TaxonomyPicker} from '@pnp/spfx-controls-react/lib/TaxonomyPicker';
import { Button } from 'office-ui-fabric-react';
import { SPOperations } from '../../Services/SPServices';

export default class Pnpcontrols extends React.Component<IPnpcontrolsProps, IPnpcontrolsState, {}> {
  private _spService:SPOperations;
  constructor(props:IPnpcontrolsProps){
    super(props);
      this.state={
        user:[],
        singleValue:[],
        multiValue:[]
      };
      this._spService=new SPOperations(this.props.context);
      this.getPeoplePicker= this.getPeoplePicker.bind(this);
      this.getSingleTaxValue=this.getSingleTaxValue.bind(this);
      this.getMultiTaxValue=this.getMultiTaxValue.bind(this);
  }
  public render(): React.ReactElement<IPnpcontrolsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.pnpcontrols} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <div>Demo of PnP People Picker Control</div>
          <PeoplePicker context={this.props.context} 
          titleText={'Employee Name'}
          placeholder='Enter your name'
          personSelectionLimit={2}
          onChange={this.getPeoplePicker}
          ensureUser={true}
          ></PeoplePicker>
          <Button text={"Submit"} onClick={()=> this._spService.SetPeoplePicker(this.props.context,"CustomList", this.state.user)}></Button>
        </div>
        <br/><br/>
        <div>
        <div>Demo of PnP Taxonomy Picker Control</div>
        <TaxonomyPicker label={"Single Value Taxonomy Control"}
          panelTitle={"Select Term"}
          context={this.props.context}
          termsetNameOrID='taxcontrol'
          allowMultipleSelections={false}
          onChange={this.getSingleTaxValue}
          ></TaxonomyPicker>
        <br/>
        <TaxonomyPicker label={"Multiple Value Taxonomy Control"}
          panelTitle={"Select Term"}
          context={this.props.context}
          termsetNameOrID='taxcontrol'
          allowMultipleSelections={true}
          onChange={this.getMultiTaxValue}>
        </TaxonomyPicker>

      <Button label='SubmitTaxonomy' text={"SubmitTaxonomy"}  
          onClick={()=>this._spService.SetTaxonomyControlValue(this.props.context,"CustomList", this.state.singleValue, this.state.multiValue)}></Button>
        </div>
      </section>
    );
  }

  public getPeoplePicker(items: any[]){
    console.log(items);
    let tempUser:any[]=[];
    items.map((item)=>{
      tempUser.push(item.id);
    })
    this.setState({user:tempUser});
  }

  public getSingleTaxValue(selectedTerm:IPickerTerms){
    this.setState({singleValue:selectedTerm})
  }

  public getMultiTaxValue(selectedTerms:IPickerTerms){
    this.setState({multiValue:selectedTerms})
  }
}
