/* eslint-disable no-undef */
import * as React from "react";
import Progress from "./Progress";
import OpenAI from "openai";
import { generate } from "../ai";
import { Button, Textarea, Input, Label } from "@fluentui/react-components";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  generatedText: string;
  startText: string;
  finalMailText: string;
  isLoading: boolean;
  baseUrl: string;
  isGenerateBusinessMailActive: boolean;
  isSummarizeMailActive: boolean;
  summary: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props) {
    super(props);

    let isGenerateBusinessMailActive;
    let isSummarizeMailActive;

    //get the current URL
    const url = window.location.href;
    console.log("URL: " + url);
    //check if the URL contains the parameter "generate"
    if (url.indexOf("compose") > -1) {
      console.log("Action: generate business mail");
      isGenerateBusinessMailActive = true;
      isSummarizeMailActive = false;
    }
    //check if the URL contains the parameter "summarize"
    if (url.indexOf("summary") > -1) {
      console.log("Action: summarize mail");
      isGenerateBusinessMailActive = false;
      isSummarizeMailActive = true;
    }

    this.state = {
      generatedText: "",
      startText: "",
      finalMailText: "",
      baseUrl: localStorage.getItem("modelBaseUrl") ?? "",
      isLoading: false,
      isGenerateBusinessMailActive: isGenerateBusinessMailActive,
      isSummarizeMailActive: isSummarizeMailActive,
      summary: "",
    };
  }

  showGenerateBusinessMail = () => {
    this.setState({ isGenerateBusinessMailActive: true, isSummarizeMailActive: false });
  };

  showSummarizeMail = () => {
    this.setState({ isGenerateBusinessMailActive: false, isSummarizeMailActive: true });
  };

  generateText = async () => {
    // eslint-disable-next-line no-undef
    var current = this;

    try {
      current.setState({ isLoading: true });
      const text = await generate(this.state.baseUrl, [
        {
          role: "system",
          content: "You are a helpful assistant that can help users to create professional business content.",
        },
        {
          role: "user",
          content: `Turn the following text into a professional business mail: ${this.state.startText}`,
        },
      ]);

      current.setState({ isLoading: false });
      current.setState({ generatedText: text });
    } catch (error) {
      console.error(error);
    }
  };

  insertIntoMail = () => {
    const finalText = this.state.finalMailText.length === 0 ? this.state.generatedText : this.state.finalMailText;

    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, {
      coercionType: Office.CoercionType.Text,
    });
  };

  onSummarize = async () => {
    try {
      this.setState({ isLoading: true });
      const summary = await this.summarizeMail();
      this.setState({ summary: summary, isLoading: false });
    } catch (error) {
      this.setState({ summary: error, isLoading: false });
    }
  };

  summarizeMail(): Promise<any> {
    const baseUrl = this.state.baseUrl;

    return new Office.Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
          //take only the first 800 words of the mail
          const mailText = asyncResult.value.split(" ").slice(0, 800).join(" ");

          //create the request body
          const text = await generate(baseUrl, [
            {
              role: "system",
              content:
                "You are a helpful assistant that can help users to better manage emails. The mail thread can be made by multiple prompts.",
            },
            {
              role: "user",
              content: `Summarize the following mail thread and summarize it with a bullet list: ${mailText}`,
            },
          ]);

          resolve(text);
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  onBaseUrlChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const newBaseUrl = event.target.value;
    this.setState({ baseUrl: newBaseUrl });
    localStorage.setItem("modelBaseUrl", newBaseUrl);
  };

  ProgressSection = () => {
    if (this.state.isLoading) {
      return <Progress title="Loading..." message="The AI is working..." />;
    } else {
      return <> </>;
    }
  };

  BusinessMailSection = () => {
    if (this.state.isGenerateBusinessMailActive) {
      return (
        <>
          <p>Briefly describe what you want to communicate in the mail:</p>
          <Textarea
            className="ms-welcome"
            onChange={(e) => this.setState({ startText: e.target.value })}
            rows={5}
            cols={40}
          />
          <p>
            <Button className="ms-welcome__action" onClick={this.generateText}>
              Generate texts
            </Button>
          </p>
          <this.ProgressSection />
          <Textarea
            className="ms-welcome"
            value={this.state.generatedText}
            onChange={(e) => this.setState({ finalMailText: e.target.value })}
            rows={15}
            cols={40}
          />
          <p>
            <Button className="ms-welcome__action" onClick={this.insertIntoMail}>
              Insert into mail
            </Button>
          </p>
        </>
      );
    } else {
      return <div> </div>;
    }
  };

  SummarizeMailSection = () => {
    if (this.state.isSummarizeMailActive) {
      return (
        <>
          <p>Summarize mail</p>
          <Button className="ms-welcome__action" onClick={this.onSummarize}>
            Summarize mail
          </Button>
          <this.ProgressSection />
          <Textarea className="ms-welcome" value={this.state.summary} rows={15} cols={40} />
        </>
      );
    } else {
      return <div> </div>;
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <main className="ms-welcome__main">
          <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            Outlook AI Assistant
          </h2>
          <Label htmlFor="baseUrl">Model Base Url</Label>
          <Input id="baseUrl" value={this.state.baseUrl} onChange={this.onBaseUrlChange} />
          <p className="ms-font-l ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            Choose your service:
          </p>
          <p>
            <Button
              className="ms-welcome__action"
              // iconProps={{ iconName: "ChevronRight" }}
              onClick={this.showGenerateBusinessMail}
            >
              Generate business mail
            </Button>
          </p>
          <p>
            <Button
              className="ms-welcome__action"
              // iconProps={{ iconName: "ChevronRight" }}
              onClick={this.showSummarizeMail}
            >
              Summarize mail
            </Button>
          </p>
          <div>
            <this.BusinessMailSection />
          </div>
          <div>
            <this.SummarizeMailSection />
          </div>
        </main>
      </div>
    );
  }
}
