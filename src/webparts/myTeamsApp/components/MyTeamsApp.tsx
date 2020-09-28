import * as React from 'react';
import styles from './MyTeamsApp.module.scss';
import { IMyTeamsAppProps } from './IMyTeamsAppProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class MyTeamsApp extends React.Component<IMyTeamsAppProps, {}> {
   
 
  public render(): React.ReactElement<IMyTeamsAppProps> {
    const teamsCtx=this.props.teamsContext.context;
    return (
     
      <div className={ styles.myTeamsApp }>
           <p className={ styles.description }>{escape(this.props.description)}</p>
           {this.props.teamsContext &&

              <table>
              <tr key={"header"}>             
                  <th>{`teamName`}</th>
                  <th>{`channelName`}</th>
                  <th>{`tenantSKU`}</th>
                  <th>{`userPrincipalName`}</th>
              
              </tr>
           <tr>
              <td>{teamsCtx.teamName}</td>
                  <td>{teamsCtx.channelName}</td>
                  <td>{teamsCtx.tenantSKU}</td>
                  <td>{teamsCtx.userPrincipalName}</td>              
           </tr>
        
            
            
            </table>
           }
      </div>
    );
  }
}
