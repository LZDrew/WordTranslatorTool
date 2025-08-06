-----

### Word Document Translation Tool

A C\# desktop application that leverages the OpenAI API for efficient and accurate batch translation of technical documents in `.docx` format.

-----

### Features

  * **Batch Translation**: Processes multiple paragraphs in a single API call to improve efficiency.
  * **Text Filtering**: Automatically skips content that doesn't need translation, such as code snippets, URLs, or Word's internal fields.
  * **Supports multiple language**: Supports multiple language pairs (e.g., Chinese, English, Japanese, Korean) and allows you to choose between GPT-3.5 and GPT-4o models.
  * **Two Translation Modes**: Offers "Keep Original + Translate" and "Replace Original (Full Translation)" options.

-----

### Getting Started

#### Prerequisites

  * **.NET Framework 4.7.2 or later**: Required to build and run the application.
  * **OpenAI API Key**: You'll need an API key to use the translation service.

#### Build and Run the Project

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/LZDrew/WordTranslatorTool.git
    ```
2.  **Open in Visual Studio**: Open the `WordTranslatorTool.sln` solution file with Visual Studio.
3.  **Build the project**: Build the solution to compile the application.
4.  **Run the application**: Run the project, and the main form will open.

-----

### Technical Details

  * **Language**: C\#
  * **Framework**: .NET Framework (Windows Forms)
  * **Document Handling**: `DocumentFormat.OpenXml` is used to read and modify `.docx` files.
  * **API Communication**: Asynchronous calls to the OpenAI Chat Completions API are made using `HttpClient`.
  * **Security**: `SecureString` ensures that the API key is handled securely.

-----

### Contributing

Contributions are welcome\! If you find a bug or have a feature request, please feel free to:

1.  Open an [Issue](https://github.com/LZDrew/WordTranslatorTool/issues).
2.  [Fork](https://github.com/LZDrew/WordTranslatorTool/fork) the repository, submit your changes, and send a Pull Request.

-----

### License

This project is licensed under the **MIT License**. See the [LICENSE](https://github.com/LZDrew/WordTranslatorTool/blob/master/LICENSE) file for details.

-----

### Contact

  * **Author**: LZDrew
  * **Email**: ahdrew51a@gmail.com
  * **GitHub**: [@LZDrew](https://github.com/LZDrew)
