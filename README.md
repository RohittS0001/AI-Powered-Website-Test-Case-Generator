# AI-Powered Website Test Case Generator

An AI-powered QA automation tool that explores a web application, captures its visible interface, and generates structured software test cases using Large Language Models (LLMs). The generated test cases can be exported to Excel for manual testing, QA documentation, and test planning.

## Overview

Manually creating comprehensive test cases for complex web applications can be time-consuming. This project helps streamline the process by combining browser automation with AI-generated test scenarios.

The tool connects to an existing browser session, captures information about accessible pages and UI elements, and uses an LLM to generate organized test cases. It supports saving progress so that test generation can continue across multiple sessions.

## Features

- **Browser interaction:** Uses Playwright to connect to a running Chromium-based browser through Chrome DevTools Protocol (CDP).
- **UI discovery:** Captures visible page content and interface elements from a logged-in web application.
- **AI-powered generation:** Uses an LLM API to generate relevant test scenarios based on captured UI information.
- **Structured test cases:** Organizes generated cases into a consistent format.
- **Excel export:** Creates formatted `.xlsx` workbooks using OpenPyXL.
- **Progress persistence:** Saves generation progress in JSON files to support continuation between sessions.
- **Duplicate reduction:** Uses normalized test-case information to identify and reduce duplicate scenarios.
- **Batch processing:** Generates test cases in manageable batches.
- **Manual control:** Allows the user to generate cases, skip a feature, or exit while saving progress.

## Technology Stack

| Technology | Purpose |
|---|---|
| Python | Main application logic |
| Playwright | Browser automation and UI inspection |
| Gemini API / Groq API | AI-powered test-case generation, depending on the configured provider |
| Requests | HTTP communication with the AI API |
| OpenPyXL | Excel workbook creation and formatting |
| Python-dotenv | Loading configuration and API keys from `.env` |
| JSON | Storing captured information and generation progress |
| Regular Expressions | Text normalization and duplicate detection |

## How It Works

1. **Browser setup:** Launches a Chromium-based browser with remote debugging enabled.
2. **Application access:** The tester logs in to the target web application manually.
3. **UI inspection:** Playwright connects to the active browser and captures information from the selected page.
4. **AI processing:** Captured interface information is sent to the configured LLM API to generate test cases.
5. **Normalization:** Generated cases are cleaned, standardized, and checked for duplicates.
6. **Progress saving:** The current results and progress are stored in JSON files.
7. **Excel generation:** The collected test cases are exported to a formatted Excel workbook for QA execution.

## Generated Test Case Information

Depending on the current implementation and configured AI prompt, test cases can include fields such as:

- Test Case ID
- Module
- Parent Feature
- Child Feature / Sub-feature
- Test Scenario
- Test Steps
- Expected Result
- Actual Result
- Execution Status
- Severity
- Priority
- Bug ID
- Comments

The exact columns depend on the workbook schema configured in the code.

## Prerequisites

Before running the project, install or prepare:

- Python 3.10 or a compatible version supported by the installed dependencies
- Google Chrome or another supported Chromium-based browser
- Git
- An API key for the AI provider configured in the project
- Access to the web application you intend to test

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/ai_qa_testcase_generator.git
cd ai_qa_testcase_generator
```

Replace `YOUR_USERNAME` with your GitHub username and use your actual repository name.

### 2. Create a virtual environment

**Windows — Git Bash:**

```bash
python -m venv .venv
source .venv/Scripts/activate
```

**Windows — Command Prompt:**

```bat
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install dependencies

If the repository contains `requirements.txt`:

```bash
pip install -r requirements.txt
```

Otherwise, install the required packages:

```bash
pip install requests python-dotenv playwright openpyxl
```

If your implementation uses the Groq Python SDK, install it as well:

```bash
pip install groq
```

### 4. Configure environment variables

Create a `.env` file in the project root.

For a Gemini-based implementation:

```env
GEMINI_API_KEY=your_gemini_api_key
GEMINI_MODEL=your_supported_gemini_model
```

For a Groq-based implementation, configure the variable names and model identifier expected by your Groq integration, for example:

```env
GROQ_API_KEY=your_groq_api_key
GROQ_MODEL=your_supported_groq_model
```

**Important:** The Gemini and Groq examples are alternatives. Use the configuration that matches your actual code. Installing the Groq package alone does not switch the application from Gemini to Groq; the API request code must also support Groq.

Never commit API keys to GitHub.

### 5. Configure Chrome remote debugging

The application connects to a running browser through Chrome DevTools Protocol, commonly at:

```text
http://127.0.0.1:9222
```

Launch Chrome with remote debugging enabled. For example, on Windows:

```bat
start chrome --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-test-profile"
```

If this command does not launch the intended Chrome installation, use the full path to your Chrome executable.

Log in to the target application in that browser before running the generator. Use a dedicated browser profile rather than your everyday profile.

### 6. Run the application

From the project root, with the virtual environment activated:

```bash
python main.py
```

Follow the terminal prompts to inspect pages and generate test cases.

## Interactive Commands

If your current version supports these commands:

| Command | Purpose |
|---|---|
| `GEN` | Generate test cases for the current feature |
| `DONE` | Skip the current feature and continue |
| `QUIT` | Save progress and exit the application |

Enter commands exactly as prompted by the running program.

## Output Files

The application may create an `output/` directory containing files such as:

```text
output/
├── Manager_Role_300_Test_Cases.xlsx
├── manager_test_progress.json
└── manual_discovery_inventory.json
```

These filenames reflect the current configuration and may differ in other versions.

- **Excel workbook:** Contains the generated test cases.
- **Progress JSON:** Stores accumulated cases and progress information for resuming work.
- **Discovery inventory:** Stores captured information about inspected pages and UI elements.

Keep the progress JSON if you want to resume a session. The workbook is generated from the saved data, so back up both the workbook and progress files before making major changes.

## Safety and Responsible Use

- Only test websites that you own or are explicitly authorized to test.
- Use a dedicated test account and browser profile.
- Avoid submitting destructive actions, deleting records, or changing production data unless explicitly authorized and safely controlled.
- Review AI-generated test cases before executing them.
- Do not include passwords, session cookies, API keys, or confidential application data in your public repository.
- Review the data sent to external AI providers before using the tool with confidential business applications.

## Limitations

- AI-generated test cases may contain inaccurate assumptions and require human review.
- Coverage depends on the pages and UI elements that the tool can inspect.
- Authentication, dynamic content, nested navigation, and browser configuration may affect discovery.
- Generating test cases through an LLM API may incur usage limits or costs.
- This tool generates QA documentation; it does not automatically guarantee that every generated test case has been executed or validated.

## Future Improvements

Potential enhancements include:

- Automatic discovery of parent, child, and nested features
- Improved form-field and validation coverage
- Integration with Jira for defect tracking
- Support for additional LLM providers
- Test-case prioritization based on risk and severity
- Integration with automated execution frameworks
- Reporting on coverage, duplicates, and generation quality

## Contributing

Contributions, bug reports, and suggestions are welcome. Fork the repository, create a feature branch, and submit a pull request describing your changes.

## License

Choose a license before distributing the project publicly. For example, the MIT License permits broad reuse subject to its terms. Add a `LICENSE` file if you choose MIT.

---

**Disclaimer:** This is an AI-assisted QA test-case generation tool. Generated scenarios should be reviewed by a qualified tester before use.
