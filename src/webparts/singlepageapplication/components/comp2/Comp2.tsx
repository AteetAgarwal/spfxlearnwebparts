import * as React from 'react';
import styles from './Comp2.module.scss';
import { IComp2Props } from './IComp2Props';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Comp2 extends React.Component<IComp2Props, {}> {
  public render(): React.ReactElement<IComp2Props> {
    const {
      description,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.singlepageapplication} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework! Component-2</h3>
        </div>
      </section>
    );
  }
}
