import * as React from 'react';
import styles from './SpfxHttpClientDemo.module.scss';
import { ISpfxHttpClientDemoProps } from './ISpfxHttpClientDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpfxHttpClientDemo extends React.Component<ISpfxHttpClientDemoProps, {}> {
  public render(): React.ReactElement<ISpfxHttpClientDemoProps> {
    const {
      // description,
      //! Lsn 4.3.7 Locate the render() method and replace it with the following code. This will create a list displaying the data contained in the spListItems property. Also notice that there's a button that has an onClick handler wired up to it.
      spListItems,
      onGetListItems,

      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.spfxHttpClientDemo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          {/* <div>Web part property value: <strong>{escape(description)}</strong></div> */}
        </div>

        <div className={styles.buttons}>
          <button type="button" onClick={this.onGetListItemsClicked}>Get Countries</button>
        </div>
        <div>
          <ul>
            {spListItems && spListItems.map((list) =>
              <li key={list.Id}>
                <strong>Id:</strong> {list.Id}, <strong>Title:</strong> {list.Title}
              </li>
            )
            }
          </ul>
        </div>

        {/* <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
        </div> */}
      </section>
    );
  }

  //! Lsn 4.3.8 Add the following event handler to the SpFxHttpClientDemo class to handle the click event on the button. This code will prevent the default action of the <a> element from navigating away from (or refreshing) the page and call the callback set on the public property, notifying the consumer of the component an event occurred.
  private onGetListItemsClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
  
    this.props.onGetListItems();
  }
}
