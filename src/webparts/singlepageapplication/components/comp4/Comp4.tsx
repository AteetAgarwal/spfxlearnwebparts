import * as React from 'react';
import styles from './Comp4.module.scss';
import { IComp4Props } from './IComp4Props';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Comp4 extends React.Component<IComp4Props, {}> {
  public render(): React.ReactElement<IComp4Props> {
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
          <h3>Welcome to SharePoint Framework! - Component-4</h3>
        </div>
      </section>
    );
  }
}
