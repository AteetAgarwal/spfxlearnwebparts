import * as React from 'react';
import styles from './Customformwebpart.module.scss';
import { ICustomformwebpartProps } from './ICustomformwebpartProps';
import {ICustomformwebpartState} from './ICustomformwebpartState';
import { escape } from '@microsoft/sp-lodash-subset';
import {Label, TextField, ChoiceGroup, Checkbox, IChoiceGroupOption, Button, Dialog, DialogFooter} from '@fluentui/react';
import {PeoplePicker, PrincipalType} from '@pnp/spfx-controls-react/lib/PeoplePicker';
import {IPickerTerms, TaxonomyPicker} from '@pnp/spfx-controls-react/lib/TaxonomyPicker';
import { DialogType, IPersonaProps, PrimaryButton } from 'office-ui-fabric-react';
import {SPOperations} from '../../Services/SPServices';

const dialogContent={
  type:DialogType.normal, 
  title:"Message", 
  subText:"Item is created successfully",
  closeButtonAriaLabel: "Close"
};
export default class Customformwebpart extends React.Component<ICustomformwebpartProps, ICustomformwebpartState, {}> {
  private trainingType:IChoiceGroupOption[]=[{key:"yes", text:"Yes"},{key:"no", text:"No"}];
  private _spService:SPOperations;
  constructor(props:ICustomformwebpartProps){
    super(props);
    this.state={
      email:"",
      mobile:"",
      address:"",
      mgrApproval:"",
      availability:false,
      employees:[],
      courses:[],
      multicourses:[],
      hideDialog:true
    }
    this.getEmail = this.getEmail.bind(this);
    this.getMobile = this.getMobile.bind(this);
    this.getAddress = this.getAddress.bind(this);
    this.getMgrApproval = this.getMgrApproval.bind(this);
    this.getAvailability = this.getAvailability.bind(this);
    this.getEmployees = this.getEmployees.bind(this);
    this.getCourse = this.getCourse.bind(this);
    this.getMultiCourse = this.getMultiCourse.bind(this);
    this.SubmitData = this.SubmitData.bind(this);
    this.Cancel = this.Cancel.bind(this);
    this._spService = new SPOperations(this.props.context);
  }
  public render(): React.ReactElement<ICustomformwebpartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.customformwebpart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}! {description}</h2>
          <div>{environmentMessage}</div>
        </div>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.title}>Training Request Form</div>
            <br/>
            <div id="dv_customForm">
              <div className={styles.grid}>
                  <div className={styles.gridRow}>
                    <div className={styles.smallCol}>
                      <Label>Employee Name<span className={styles['validation-mandatory']}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <PeoplePicker 
                          context={this.props.context}
                          placeholder='Enter your name'
                          ensureUser={true}
                          personSelectionLimit={3}
                          groupName={""} //Leave this blank in case you want to filter from all users
                          showtooltip={false}
                          disabled={false}
                          showHiddenInUI={false}
                          resolveDelay={1000}
                          principalTypes={[PrincipalType.User]}
                          onChange={this.getEmployees}
                          defaultSelectedUsers={this.state.employees}
                      ></PeoplePicker>
                      <div id="validation_employee" className={styles['form-validation']}>You can't leave this blank.</div>
                    </div>
                  </div>
                  <div className={styles.gridRow}>
                    <div className={styles.smallCol}>
                      <Label>Email<span className={styles['validation-mandatory']}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <TextField placeholder='Enter your email here' onChange={this.getEmail} value={this.state.email}></TextField>
                      <div id="validation_email" className={styles['form-validation']}>You can't leave this blank.</div>
                    </div>
                  </div>
                  <div className={styles.gridRow}>
                    <div className={styles.smallCol}>
                      <Label>Mobile<span className={styles['validation-mandatory']}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <TextField type='Number' placeholder='Enter your 10 digit mobile number' value={this.state.mobile} onChange={this.getMobile}></TextField>
                      <div id="validation_mobile" className={styles['form-validation']}>You can't leave this blank.</div>
                    </div>
                  </div>
                  <div className={styles.gridRow}>
                    <div className={styles.smallCol}>
                      <Label>Address</Label>
                    </div>
                    <div className={styles.largeCol}>
                      <TextField multiline={true} placeholder='Enter your address' value={this.state.address} onChange={this.getAddress}></TextField>
                    </div>
                  </div>
                  <div className={styles.gridRow}>
                    <div className={styles.smallCol}>
                      <Label>Choose Your Course<span className={styles['validation-mandatory']}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <TaxonomyPicker
                          context={this.props.context}
                          label=''
                          panelTitle='Select Term'
                          termsetNameOrID='Skills'
                          placeholder='Select Course'
                          isTermSetSelectable={false}
                          onChange={this.getCourse}
                          initialValues={this.state.courses}
                      ></TaxonomyPicker>
                      <div id="validation_course" className={styles['form-validation']}>You can't leave this blank.</div>
                    </div>
                  </div>
                  <div className={styles.gridRow}>
                    <div className={styles.smallCol}>
                      <Label>Choose Mulitple Courses<span className={styles['validation-mandatory']}>*</span></Label>
                    </div>
                    <div className={styles.largeCol}>
                      <TaxonomyPicker
                          context={this.props.context}
                          label=''
                          panelTitle='Select Term'
                          termsetNameOrID='Skills'
                          allowMultipleSelections={true}
                          placeholder='Select Course'
                          isTermSetSelectable={false}
                          onChange={this.getMultiCourse}
                          initialValues={this.state.multicourses}
                      ></TaxonomyPicker>
                      <div id="validation_multicourse" className={styles['form-validation']}>You can't leave this blank.</div>
                    </div>
                  </div>
                  <div className={styles.gridRow}>
                    <div className={styles.smallCol}>
                      <Label>Do you have manager approval?</Label>
                    </div>
                    <div className={styles.largeCol}>
                      <ChoiceGroup options={this.trainingType} defaultSelectedKey={this.state.mgrApproval} selectedKey={this.state.mgrApproval}
                       onChange={this.getMgrApproval}></ChoiceGroup>
                    </div>
                  </div>
                  <div className={styles.gridRow}>
                    <div className={styles.smallCol}>
                      <Label>Available On Weekdays</Label>
                    </div>
                    <div className={styles.largeCol}>
                      <Checkbox label='Yes' checked={this.state.availability} onChange={this.getAvailability}></Checkbox>
                    </div>
                  </div>
                  <div className={styles.gridRow}>
                    <div className={styles.largeCol}>
                      <Button className={styles.button} text="Submit" onClick={this.SubmitData}></Button>
                      <Button className={styles.button} text="Cancel" onClick={this.Cancel}></Button>
                    </div>
                    <Dialog 
                        onDismiss={this.toggleDialog}
                        dialogContentProps={dialogContent}
                        hidden={this.state.hideDialog}
                    >
                      <DialogFooter>
                        < PrimaryButton text="Close" onClick={this.toggleDialog}></PrimaryButton>
                      </DialogFooter>
                    </Dialog>
                  </div>
              </div>
            </div>
          </div>
        </div>
      </section>
    );
  }

  public SubmitData(){
    let validation:boolean=true;
    if(this.state.email == null || this.state.email == undefined || this.state.email == ""){
      validation=false;
      document.getElementById('validation_email').setAttribute("style","display:block !important");
    }
    if(this.state.mobile == null || this.state.mobile == undefined || this.state.mobile == ""){
      validation=false;
      document.getElementById('validation_mobile').setAttribute("style","display:block !important");
    }
    if(this.state.employees == null || this.state.employees == undefined || this.state.employees.length == 0){
      validation=false;
      document.getElementById('validation_employee').setAttribute("style","display:block !important");
    }
    if(this.state.courses == null || this.state.courses == undefined || this.state.courses.length == 0){
      validation=false;
      document.getElementById('validation_course').setAttribute("style","display:block !important");
    }
    if(this.state.multicourses == null || this.state.multicourses == undefined || this.state.multicourses.length == 0){
      validation=false;
      document.getElementById('validation_multicourse').setAttribute("style","display:block !important");
    }
    if(validation){
      let multicoursesVal:string="";
      this.state.multicourses.map((course)=>{
        multicoursesVal+=`-1;#${course.name}|${course.key};#`
      })
      this._spService.SubmitFormData("EmployeesData","MultiCourse_0",multicoursesVal,this.state).then((response)=>{
        console.log(response);
        this.setState({hideDialog:false});
      })
    }
  }

  public Cancel(){
    this.setState({
      email:"",
      mobile:"",
      address:"",
      mgrApproval:"",
      availability:false,
      employees:[],
      courses:[],
      multicourses:[]
    });
  }

  private toggleDialog = (event:any)=>{
    this.setState({hideDialog:!this.state.hideDialog});
  }

  private getEmail(event:any,val:string){
    this.setState({email:val});
  }

  private getMobile(event:any,val:string){
    this.setState({mobile:val});
  }

  private getAddress(event:any,val:string){
    this.setState({address:val});
  }

  private getMgrApproval(event:any,val:IChoiceGroupOption){
    this.setState({mgrApproval:val.key});
  }

  private getAvailability(event:any,val:boolean){
    this.setState({availability:val});
  }

  private getEmployees(val:IPersonaProps[]){
    let emp:any[]=[];
    val.map((item)=>{
      emp.push(item.id);
    })
    this.setState({employees:emp});
  }

  private getCourse(val:IPickerTerms){
    this.setState({courses:val});
  }

  private getMultiCourse(val:IPickerTerms){
    this.setState({multicourses:val});
  }
}
