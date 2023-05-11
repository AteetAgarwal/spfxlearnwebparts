import * as React from 'react';
import styles from './Responsivegrid.module.scss';
import { IResponsivegridProps } from './IResponsivegridProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {TextField, Button} from "office-ui-fabric-react"

export default class Responsivegrid extends React.Component<IResponsivegridProps, {}> {
  public render(): React.ReactElement<IResponsivegridProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.responsivegrid} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div id="dv_custom" className={styles['ms-Fabric']} dir="ltr">
          <div className={styles['ms-Grid']}>
            <div className={styles['ms-Grid-row']}>
            <div className={styles['ms-Grid-col'] + ' ' +  styles['ms-sm6'] + ' '+  styles['ms-md6'] + ' ' + styles['ms-lg2']}>Fluent UI Text</div>
              <div className={styles['ms-Grid-col'] + ' ' +  styles['ms-sm6'] + ' '+  styles['ms-md6'] + ' ' + styles['ms-lg10']}>
                <TextField></TextField>
              </div>
            </div>
            <div className={styles['ms-Grid-row']}>
              <div className={styles['ms-Grid-col'] + ' ' +  styles['ms-sm6'] + ' '+  styles['ms-md6'] + ' ' + styles['ms-lg2']}>FluentUI Button</div>
              <div className={styles['ms-Grid-col'] + ' ' +  styles['ms-sm6'] + ' '+  styles['ms-md6'] + ' ' + styles['ms-lg10']}>
                <Button></Button>
                </div>
            </div>
          </div>
        </div>
      </section>
    );
  }
}
