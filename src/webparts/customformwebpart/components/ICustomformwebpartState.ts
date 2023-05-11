import {IPickerTerms} from '@pnp/spfx-controls-react/lib/TaxonomyPicker';
export interface ICustomformwebpartState {
  email:string;
  mobile:string;
  address:string;
  mgrApproval:string;
  availability:boolean;
  employees: any[];
  courses: IPickerTerms;
  multicourses: IPickerTerms;
  hideDialog:boolean;
}
