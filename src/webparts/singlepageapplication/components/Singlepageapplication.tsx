import * as React from 'react';
//import styles from './Singlepageapplication.module.scss';
import { ISinglepageapplicationProps } from './ISinglepageapplicationProps';
import Comp1 from './navigationcomp1/navigationcomp1';
import Comp4 from './comp4/Comp4';
import Comp3 from './comp3/Comp3';
import Comp2 from './comp2/Comp2';
import { HashRouter, Navigate, Route, Routes } from 'react-router-dom';
import { Stack, StackItem } from 'office-ui-fabric-react';
//import { escape } from '@microsoft/sp-lodash-subset';

export default class Singlepageapplication extends React.Component<ISinglepageapplicationProps, {}> {
  public render(): React.ReactElement<ISinglepageapplicationProps> {
    const {
      description,
      userDisplayName
    } = this.props;
    return (
      <HashRouter>
        <Stack>
          <Comp1></Comp1>
          <StackItem>
            <switch>
              <Routes>
                <Route path="/comp2"
                Component={()=><Comp2 userDisplayName={userDisplayName} description={description}/>}></Route>
                <Route path="/comp3"
                Component={()=><Comp3 userDisplayName={userDisplayName} description={description}/>}></Route>
                <Route path="/comp4"
                Component={()=><Comp4 userDisplayName={userDisplayName} description={description}/>}></Route>
                <Route path="*" element={<Navigate to="/comp2" replace />}></Route>
              </Routes>
            </switch>
          </StackItem>
        </Stack>
      </HashRouter>
    );
  }
}
