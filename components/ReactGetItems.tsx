import * as React from 'react';
import styles from './ReactGetItems.module.scss';
import { IReactGetItemsProps } from './IReactGetItemsProps';
import { IReactGetItemsState, IListItem } from './IReactGetItemsState';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { TextField, DefaultButton, PrimaryButton, Stack, IStackTokens, IIconProps } from 'office-ui-fabric-react/lib/';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jquery from 'jquery';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
const stackTokens: IStackTokens = { childrenGap: 40 };

export interface IReactGetItemsState{    
  items:[    
        {    
          "EmpID": "",    
          "StaffHour": "", 
          "LoginHour": "", 
          "FirstName ": "", 
          "CreatedBy": "",    
          "CreatedDate":"",    
          "Description":""  
        }]    
} 

export default class ReactGetItems extends React.Component<IReactGetItemsProps, IReactGetItemsState> {

   constructor(props: IReactGetItemsProps, state: IReactGetItemsState){    
    super(props);    
    this.state = {    
      items: [    
        {    
          "EmpID": "",    
          "StaffHour": "", 
          "LoginHour": "", 
          "FirstName ": "", 
          "CreatedBy": "",    
          "CreatedDate":"",    
          "Description":"" 
        }    
      ]    
    };    
    sp.setup({
      spfxContext: this.props.spcontext
    });
    if (Environment.type === EnvironmentType.SharePoint) {
      this._getListItems();
    }
    else if (Environment.type === EnvironmentType.Local) {
      // return (<div>Whoops! you are using local host...</div>);
    }
  } 

  async _getListItems() {
    const allItems: any[] = await sp.web.lists.getByTitle("Employeeee").items.getAll();
    console.log(allItems);
    let items: IListItem[] = [];
    allItems.forEach(element => {
      items.push({ EmpID: element.EmpID, StaffHour: element.StaffHour, LoginHour: element.LoginHour, FirstName: element.FirstName, CreatedBy:element.CreatedBy, CreatedDate: element.CreatedDate, Description: element.Description});
    });
    //this.setState({updateText: items });
  }

  public render(): React.ReactElement<IReactGetItemsProps> {
    return (
      
      <div className={ styles.reactGetItems }>        
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
            <div className={styles.headerStyle} >    
            </div>  
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <div className={styles.CellStyle}>EmplID</div>    
              <div className={styles.CellStyle}>StaffHour </div>    
              <div className={styles.CellStyle}>LoginHour</div>    
              <div className={styles.CellStyle}>CreatedDate</div>  
              <div className={styles.CellStyle}>CreatedBy</div> 
              <div className={styles.CellStyle}>Description</div> 
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
};


