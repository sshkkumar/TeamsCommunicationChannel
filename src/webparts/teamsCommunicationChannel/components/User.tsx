import * as React from "react";
import { render } from "react-dom";
import Styles from "./Styles";
import compStyles from "./TeamsCommunicationChannel.module.scss";
import "./App.css";
import Dropzone from "react-dropzone";
import "bootstrap/dist/css/bootstrap.css";
import styled, { css } from "styled-components";

import { IUserProps } from "./IUserProps";
import { IUserState } from "./IUserState";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration,
  ISPHttpClientOptions
} from "@microsoft/sp-http";
import { sp, ItemAddResult } from "@pnp/sp";

import { TextField } from "office-ui-fabric-react/lib/TextField";
import {
  PrimaryButton,
  DefaultButton
} from "office-ui-fabric-react/lib/components/Button";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { any } from "prop-types";

import * as jQuery from "jquery";
import { SPComponentLoader } from "@microsoft/sp-loader";

export interface IDocumentUploadResult {
  Name: string;
  ServerRelativeUrl: string;
}

export interface IDocumentItem {
  ID: string;
}

class User extends React.Component<IUserProps, IUserState> {
  constructor(props) {
    super(props);
    // this.handleTitle = this.handleTitle.bind(this);
    // this.handleDesc = this.handleDesc.bind(this);
    // this._onCheckboxChange = this._onCheckboxChange.bind(this);
    // this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    // this.createItem = this.createItem.bind(this);
    // this.onTaxPickerChange = this.onTaxPickerChange.bind(this);
    // this._getManager = this._getManager.bind(this);
    this.state = {
      title: "",
      description: "",
      dpselectedItem: undefined,
      dpselectedItems: [],
      userManagerIDs: [],
      customerName: "",
      documentType: "",
      required: "This is required",
      onSubmission: false,
      filesToUpload: [],
      items: [],
      users: [],
      context: this.props.spContext,
      teamsContext: this.props.teamsContext,
      message: ""
    };
  }
  public uploadFile(context, teamsContext) {
    var files = (document.getElementById("inputFile") as HTMLInputElement)
      .files;

    var file = files[0];
    if (file != undefined || file != null) {
      let spOpts: ISPHttpClientOptions = {
        headers: new Headers(),
        method: "POST",
        mode: "cors",
        // headers: {
        //   // Accept: "application/json",
        //   // "Content-Type": "application/json" //,
        //   //"X-RequestDigest": this.props.digest
        // },
        body: file
      };
      //${teamsContext.teamSiteUrl}
      var url = `https://m365x846523.sharepoint.com/sites/Dep1/_api/Web/Lists/getByTitle('Documents')/rootfolder/folders('Outgoing')/files/add(overwrite=true,url='${file.name}')`;

      return context.spHttpClient
        .post(url, SPHttpClient.configurations.v1, spOpts)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    }
  }

  public onSubmit(context, teamsContext, formDigest) {
    let title = this.state.title;
    let description = this.state.description;
    let documentType = this.state.documentType;

    this.uploadFile(context, teamsContext) //.then(res => {
      //.then(this.getItemFromFile)
      .then(fileResponse => {
        let serverRelativeUrl = fileResponse.ServerRelativeUrl;
        //https://m365x846523.sharepoint.com/sites/Dep1
        jQuery.ajax({
          url: `https://m365x846523.sharepoint.com/sites/Dep1/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')/ListItemAllFields`,
          type: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
            "X-RequestDigest": formDigest,
            crossDomain: "true",
            credentials: "include"
          },
          xhrFields: { withCredentials: true },
          // data:
          //   "{'__metadata': { 'type': 'SP.Data.Shared_x0020_DocumentsItem' },Title:'IT Change Management'}",
          data:
            "{'__metadata': { 'type': 'SP.Data.Shared_x0020_DocumentsItem' },Title:'" +
            title +
            "',TargetDepartment:'IT',DocumentType:'" +
            documentType +
            "',Description1:'" +
            description +
            "'}",
          // data: `{'__metadata': { 'type': 'SP.Data.Shared_x0020_DocumentsItem' },
          //   Title:${this.state.title},
          //   DocumentType: ${this.state.documentType},
          //   KpiDescription: ${this.state.description},
          //   TargetDepartment: 'IT'}`,
          success: function(data) {
            console.log(data.d.results);
          },
          error: function(error) {
            //alert(JSON.stringify(error));
          }
        });
      });

    //}); //Upload file ends
  }

  // onDrop = files => {
  //   var uploadedFiles = (document.getElementById(
  //     "fileUpload"
  //   ) as HTMLInputElement).files;
  //   var file = uploadedFiles[0];

  //   // console.log("files: " + files[0].name);
  //   var fileItems = files.map((item, key) => <li key={key}>{item.name}</li>);
  //   this.setState({ filesToUpload: files, items: fileItems });
  // };

  // onSubmit = async values => {
  //   window.alert(JSON.stringify(values));
  // };

  submitForm = e => {
    var form = e.target;
    e.preventDefault();
    form.reset();

    // e.preventDefault();
    // e.target.reset();
  };

  resetForm = () => {
    this.setState({
      title: "",
      description: "",
      documentType: "",
      message: "Document submitted successfully!"
    });
  };

  render() {
    let _teamsContext = this.props.teamsContext;
    let _spContext = this.props.spContext;
    let _formDigest = this.props.formDigest;

    return (
      <Styles>
        <div>
          Please fill in the form below to request for document approval.
        </div>
        {/* {e => {
            e.preventDefault();
          }} */}
        <form onSubmit={this.submitForm} onReset={this.resetForm}>
          <div className={compStyles.container}>
            <div className="form-group row">
              <div className="col-md-4">
                <label className="ms-Label">Title</label>
              </div>
              <div className="col-md-7">
                <TextField
                  placeholder="Please Enter Title"
                  required={true}
                  value={this.state.title}
                  onChanged={this.handleTitle.bind(this)}
                  errorMessage={
                    this.state.title.length === 0 &&
                    this.state.onSubmission === true
                      ? this.state.required
                      : ""
                  }
                />
              </div>

              {/*Description*/}
              <div className="col-md-4 space">
                <label className="ms-Label">Description</label>
              </div>
              <div className="col-md-7 space">
                <TextField
                  placeholder="Please Enter Description"
                  multiline={true}
                  required={true}
                  value={this.state.description}
                  onChanged={this.handleDescription.bind(this)}
                  errorMessage={
                    this.state.description.length === 0 &&
                    this.state.onSubmission === true
                      ? this.state.required
                      : ""
                  }
                />
              </div>
              {/*Customer Name*/}
              {/* <div className="col-md-4 space">
                <label className="ms-Label">Customer Name</label>
              </div>
              <div className="col-md-7 space">
                <TextField
                  placeholder="Please Enter Customer Details"
                  onChanged={this.handleCustomer.bind(this)}
                  errorMessage={
                    this.state.title.length === 0 &&
                    this.state.onSubmission === true
                      ? this.state.required
                      : ""
                  }
                />
              </div> */}
              {/*Department*/}
              <div className="col-md-4 space">
                <label className="ms-Label">Department</label>
              </div>
              <div className="col-md-7 space">
                <Dropdown
                  placeHolder="Select an Option"
                  label=""
                  id="component"
                  //onChanged={this.handleDepartmentChange.bind(this)}
                  selectedKey={
                    this.state.dpselectedItem
                      ? this.state.dpselectedItem.key
                      : undefined
                  }
                  ariaLabel="Basic dropdown example"
                  options={[
                    { key: "IT", text: "IT" },
                    { key: "Finance", text: "Finance" },
                    { key: "Employee", text: "Employee" }
                  ]}
                />
                {/* onChanged={this._changeState}
                onFocus={this._log("onFocus called")}
                onBlur={this._log("onBlur called")} */}
              </div>

              {/*Document Type*/}
              <div className="col-md-4">
                <label className="ms-Label">Document Type</label>
              </div>
              <div className="col-md-7">
                <TextField
                  placeholder="Please Enter Document Type"
                  onChanged={this.handleDocumentType.bind(this)}
                  value={this.state.documentType}
                />
              </div>

              {/*Approver*/}
              <div className="col-md-4 space">
                <label className="ms-Label">Approver</label>
              </div>
              <div className="col-md-7 space">
                <PeoplePicker
                  context={this.props.spContext} //titleText="People Picker"
                  personSelectionLimit={3}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  isRequired={true}
                  disabled={false}
                  ensureUser={true}
                  selectedItems={this.getPeoplePickerItems.bind(this)}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  errorMessage="Please enter approver"

                  // peoplePickerWPclassName="form-control"
                  // peoplePickerCntrlclassName="form-control"
                />
              </div>

              {/*Dropzone*/}
              <div className="col-md-4 space">
                <label className="ms-Label">Upload Document</label>
              </div>
              <div className="col-md-7 space">
                <input type="file" id="inputFile" />
              </div>

              {/*<div className="col-md-11">
                 <Dropzone onDrop={acceptedFiles => console.log(acceptedFiles)}>
                  {({ getRootProps, getInputProps }) => (
                    <section>
                      <div {...getRootProps()}>
                        <input {...getInputProps()} />
                        <p>
                          Drag 'n' drop some files here, or click to select
                          files
                        </p>
                      </div>
                    </section>
                  )}
                </Dropzone> */}
              {/* <Dropzone onDrop={this.onDrop}>
                  {({
                    getRootProps,
                    getInputProps,
                    isDragActive,
                    isDragAccept,
                    isDragReject,
                    acceptedFiles,
                    rejectedFiles
                  }) => {
                    return (
                      <div>
                        <div {...getRootProps()} className="dropzone">
                          <input {...getInputProps()} />
                          {isDragActive
                            ? "Drop here..."
                            : "Click here or drag a file to upload."}
                          {/* <div>
                                <h1>Drag and drop file here</h1>
                            </div>
                          {isDragReject && <div>Unsupported file type...</div>}
                        </div>
                      </div>
                    );
                  }}
                </Dropzone> */}

              {/* <Dropzone.default onDrop={this.onDrop}>
                  {({ getRootProps, getInputProps, isDragActive }) => (
                    <div {...getRootProps()} className="dropzone">
                      <input id="fileUpload" {...getInputProps()} />
                      {isDragActive
                        ? "Drop here..."
                        : "Click here or drag a file to upload."}
                    </div>
                  )}
                </Dropzone.default> */}
              {/* <Dropzone.default onDrop={this.onDrop}>
                  {({ getRootProps, getInputProps }) => (
                    <div className="container">
                      <div
                        {...getRootProps({
                          className: "dropzone",
                          onDrop: event => event.stopPropagation()
                        })}
                      >
                        <input {...getInputProps()} />
                        <p>
                          Drag 'n' drop some files here, or click to select
                          files
                        </p>
                      </div>
                    </div>
                  )}
                </Dropzone.default>
              </div>*/}

              {/*End*/}
            </div>
          </div>
          {/*Buttons */}

          <div className="buttons space">
            <button
              type="submit"
              onClick={this.onSubmit.bind(
                this,
                _spContext,
                _teamsContext,
                _formDigest
              )} //_teamsContext
            >
              Submit
            </button>
            <button type="button">Reset</button>
          </div>
          <div className="text-info">{this.state.message}</div>
        </form>
      </Styles>
    );
  }
  validateForm() {
    throw new Error("Method not implemented.");
  }

  private handleTitle(value: string): void {
    return this.setState({
      title: value
    });
  }

  private handleDescription(value: string): void {
    return this.setState({
      description: value
    });
  }

  private handleCustomer(value: string): void {
    this.setState({
      customerName: value
    });
  }

  private handleDepartmentChange(value: string): void {
    return this.setState({
      dpselectedItem: { key: value }
    });
  }

  private handleDocumentType(value: string): void {
    this.setState({
      documentType: value
    });
  }

  private getPeoplePickerItems(values: any[]) {
    console.log("Items:", values);
    // this.state.userManagerIDs.length = 0;
    // let tempuserMngrArr = [];
    // for (let item in values) {
    //   tempuserMngrArr.push(values[item].id);
    // }
    // this.setState({ userManagerIDs: tempuserMngrArr });

    // return this.setState({
    //   pplPicker: value
    // });
  }
}

export default User;
