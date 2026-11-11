Hello and a very good morning. This is a presentation about the Captable Automation Project, focusing on Phase I. In this presentation I will walk through the demo of this automation flow and I will explain how the flow works. I will explain what the efficiencies of this current automation are over all the extraction methods that have been followed across the team. I will show some edge case scenarios where both the analyst and AI models were not able to process the filings and extract the data.

Let me execute the automation and while it is running I will walk through the flow of this automation.

Whenever there is a new ticket generated, this automation will be executed. That is the trigger for it. Currently this is being built and tested with the UK jurisdiction and it is scalable to other jurisdictions as well. 

This Jira space is an exact replica of UK Captable's tickets. Whenever there is a new filing filed, automation triggers a new ticket and it will have the PDF which is newly issued. From that PDF the analyst downloads it, extracts the data, validates it, and then prepares the cap table.

This automation does everything from end to end that an analyst does and makes the Excel ready for the research, skipping downloading, extracting, cleaning, validating, and then applying formulas, pivot tables, and all the processes in Excel. 

It extracts the data directly from the PDF and based on the sort of filing it uses its extraction method. If it is a readable filing it uses a Python package. If it is a non-readable filing, a printed document, or an unstructured document, it uses Microsoft's Document Intelligence for converting the non-readable text into readable text.

And in the final line there is an AI model which is called whenever there is a discrepancy in the extracted data. Whenever an analyst extracts the data, he has to validate it against certain data points which will be programmed in the script. When it does that, when it is not able to validate it, it calls the AI and it feeds the PDF.

This AI model will find the discrepancy in where it is exactly filing, correct those mistakes, do those rectifications, and complete the Excel. Once the data is converted and output into Excel, it is uploaded into a team Excel SharePoint and that link to that Excel will be provided inside the Jira ticket 

The result for this run has been delivered and let me walk you through the run report. As you can see this has been executed within this X amount of time and this much data has been processed. This is the total time consumed. On top of it this is the token cost which is being used for it and the cost associated with using that particular AI model. 


_-----------

Now let me explain the advantages and efficiency of this automation workflow compared to the current workflow. Currently each analyst has their own way of processing the data and preparing cap table, missing sources and Excel in SharePoint, different extraction methods and validation by the previous analyst make it hard, in TLC cases, for the analyst who is working currently to validate the previously created cap table. Even cloning the cap table, for that reason, takes time. 

All the above-mentioned major points of the current workflow problems will be sorted in this automated workflow.

The ultimate goal of this project is to completely eliminate the step of downloading, extracting, validating data 

When the analyst opens the ticket, he has to have an Excel which is ready for the research

--------------

Let me show those edge cases where AI models created inside a Copilot agent were not able to process and analysts are also not able to process 

This case is one of a huge volume of data, around 350+ pages in this one confirmation statement, and this is one of the highly valued companies. This is the topmost valued company in our UK dataset and it is missing a cap table. Like this I have a few other cases as well where high-value companies were missing cap tables because of the CS issue. 

Because of this, with such a large volume of unstructured data, it is really hard for a human to manually extract it and validate it. All the current flows fail at it. The AI agent particularly fails because of the runtime and the token limit associated with it, and the limit of the prompt which can be fed to copilot agents. We cannot feed multi-line prompts. The prompt's number of words is being set. These are the barriers

And let me explain how a file, which has not been able to be processed even by AI agents with the superior model, has still been processed through this automation using a lightweight model (with lesser runtime and a highly efficient cost of tokens spent on) 

Whenever there is a filing, we currently feed the entire file to the AI agent and it has to break it down from top to bottom. It has to get all the data from it. It has to create the Excel inside the agent and it has to do all the things that happen inside the agent. 

This automation segregates everything from identifying the documentation and using the extraction method to the final step: if not able to validate, it goes for the AI model and feeds the data in a way that it is able to process.

Unlike the other models we won't be fully providing the entire 350+ pages. The API limit won't allow it so we will be feeding 20 pages. For that amount of pages the AI model is easily able to process it and store the data.

Once everything is consolidated after the final page, it extracts and validates against the statement of capital and gives the capital. 

The Excel part and the pivot table process of this current automation are not happening inside the AI agent. Once the data has been extracted it is again the job of the Python script to do all these.

This second case is of unstructured data. Usually filings are structured in a way and the AI agents are prompted to process the filing in that certain way. Over here it is such a complex structure and the data is also assigned in a way the AI agent is not able to process inside Microsoft Copilot when run against this model, which has been capable of reasoning (because I have tested it over 2,500+ filings and trained the model

And the third case is of handwritten filing, which itself, for an analyst reading with the naked eye, is a little bit trickier. These cases as well, the automation is able to process and extract and validate the data.

Cool and this is the end of this presentation. Hope I presented things in a clear way. I will send those further reports and the executive presentation with all the necessary data points by the end of this week. 

In the presentation I will be including all the necessary data points, with all the phases which have been planned part of Capable automation 

Phase two: its goal is to completely eliminate the user validating the data and extracting it, and make sure it is research-ready and covers other jurisdictions that have high volume, including complex jurisdictions such as Germany

Phase 3: Skipping the research part of the analyst.

For example when there is an individual that the analyst has been looking for, if it is available in the prequin dataset inside the portfolio section, the analyst opens the link, gets the ID, and pastes it in Excel. By connecting to the prequin database using an API, this process can be automated. This is the goal of Phase 3. 

Phase 4 would be helping an analyst with the research. If there are some data points which are missing in our frequent database, the agent is capable of doing the same way an analyst does the research and gets the data. This is phase 4 of this build. 

Phase 5 would be complete end to end, from extraction to preparation of data. The analyst only has to validate it once he approves and it has been uploaded to Cura. To make things even easier I will also plan to implement the automation which I proposed earlier: uploading the Excel directly into Cura, which will be covered in this phase. This is the final phase of this project, which is planned. 

Confidence in this project's success.
1. Currently we have unstructured data, which is extremely hard to understand and train AI models on. Fortunately by using this extraction method across the jurisdiction, we will be able to not just structure the data (which has been useful for the analyst) but also make it useful for training the AI models.
2. Central report: Whenever there has been execution we will collect those run reports and varied metadata files. How it was filed, why it was filed, and all the details will be provided in the run report. Currently as well, I've been having it, along with the tokens spent on it associated with it and all the important details of the bill, which helps identify the volume and helps train the models.