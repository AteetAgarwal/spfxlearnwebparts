import * as React from 'react';
import { Inavigationcomp1Props } from './Inavigationcomp1Props';
import {INavLinkGroup, Nav} from 'office-ui-fabric-react';
//import { NavLink } from 'react-router-dom';

const group:INavLinkGroup[]=[
  {
    links:[{name:"Component2", url:"#/comp2", key:"comp2"},
    {name:"Component3", url:"#/comp3", key:"comp3" },
    {name:"Component4", url:"#/comp4", key:"comp4" }],
  }
];
export default class navigationcomp1 extends React.Component<Inavigationcomp1Props, {}> {
  public render(): React.ReactElement<Inavigationcomp1Props> {
    return (
      <Nav initialSelectedKey='comp2'
       groups={group}>
      </Nav>
      /*<React.Fragment>
        <ul>
          <li><NavLink to="#/comp2">Component2</NavLink></li>
          <li><NavLink to="#/comp3">Component3</NavLink></li>
          <li><NavLink to="#/comp4">Component4</NavLink></li>
        </ul>
      </React.Fragment>*/
    );
  }
}
