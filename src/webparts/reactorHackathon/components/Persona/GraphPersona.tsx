import * as React from 'react';
import styles from './GraphPersona.module.scss';
import { IGraphPersonaProps } from './IGraphPersonaProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IGraphPersonaState } from './IGraphPersonaState';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {  SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http'; 

import {
  Persona,
  PersonaSize
} from 'office-ui-fabric-react/lib/components/Persona';
import { Link } from 'office-ui-fabric-react/lib/components/Link';

export default class GraphPersona extends React.Component<IGraphPersonaProps, IGraphPersonaState> {
  constructor(props: IGraphPersonaProps) {
    super(props);

    this.state = {
      name: '',
      email: '',
      phone: '',
      image: null
    };
  }

  public componentDidMount(): void {
    
    this.props.graphClient
      .api('me')
      .get((error: any, user: MicrosoftGraph.User, rawResponse?: any) => {
        this.setState({
          name: user.displayName,
          email: user.mail,
           //phone: user.businessPhones[0]
        });
      });

    let webUrl = "https://groverale.sharepoint.com/sites/YammerSentiment"
      //let requestUrl = webUrl.concat("/_api/web/Lists/GetByTitle('AllCompany')/ItemCount")   
    //let requestUrl = webUrl.concat("/_api/web/Lists/GetByTitle('AllCompany')/items")

    //Filter by LoginName (i:0#.f|membership|r@tenant-name.onmicrosoft.com)
    let userToken = `i:0#.f|membership|${this.props.spfxContext.pageContext.user.loginName}`;

    let requestUrl = `${webUrl}/_api/web/Lists/GetByTitle('AllCompany')/items?$filter=PostedBy/Name eq '${encodeURIComponent(userToken)}'`

    this.props.spfxContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)  
    .then((response: SPHttpClientResponse) => {  
        if (response.ok) {  
            response.json().then((responseJSON) => {  
                if (responseJSON!=null && responseJSON.value!=null){  
        let itemCount:number = parseInt(responseJSON.value.length.toString());
        //let itemCount:number = 0
        let totalScore:number = 0
        responseJSON.value.forEach(element => {

            totalScore += element.Score
        });
        
        let sentiment:number = totalScore / itemCount

        this.setState({ phone: "Sentiment Score: " + sentiment.toString() });
        }  
      });  
    } else {
      console.log("ERROR in request: ")
    } 
    });

    // this.props.spfxContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)  
    // .then((response: SPHttpClientResponse) => {  
    //     if (response.ok) {  
    //         response.json().then((responseJSON) => {  
    //             if (responseJSON!=null && responseJSON.value!=null){  
    //                 let items:any[] = responseJSON.value;  
    //                 this.setState({ phone: "Sentiment: " + items.length.toString() });
    //             }  
    //         });  
    //     }  
    // }); 
    
    

    this.props.graphClient
      .api('/me/photo/$value')
      .responseType('blob')
      .get((err: any, photoResponse: any, rawResponse: any) => {
        const blobUrl = window.URL.createObjectURL(photoResponse);
        this.setState({ image: blobUrl });
      });

      
  
    
  }

  private _renderMail = () => {
    if (this.state.email) {
      return <Link href={`mailto:${this.state.email}`}>{this.state.email}</Link>;
    } else {
      return <div />;
    }
  }

  private _renderPhone = () => {
    if (this.state.phone) {
      return <p>{this.state.phone}</p>;
    } else {
      return <div />;
    }
  }

  public render(): React.ReactElement<IGraphPersonaProps> {
    return (
      <Persona primaryText={this.state.name}
              secondaryText={this.state.email}
              onRenderSecondaryText={this._renderMail}
              tertiaryText={this.state.phone}
              onRenderTertiaryText={this._renderPhone}
              imageUrl={this.state.image}
              size={PersonaSize.size100} />
    );
  }
}