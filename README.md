# ComPDF Conversion SDK for Windows

As part of the KDAN ecosystem, ComPDF Conversion SDK helps you build high-quality PDF conversion features into your Windows application. Empower your apps to convert PDF files into editable Office files, images, HTML, RTF, CSV, TXT, and more — while preserving layout, tables, text structure, and formatting. Whether you're building a document processing tool for enterprise users or a file converter for consumers, you can integrate reliable PDF conversion capabilities using a straightforward C# API.

> If you find this library helpful, please consider giving us a ⭐ **Star** on GitHub! Have feedback or questions? Join the conversation in our [Discussions](https://github.com/orgs/ComPDFKit/discussions).

**Supported Features**

* Convert PDF to Word (.docx)

* Convert PDF to Excel (.xlsx)

* Convert PDF to PowerPoint (.pptx)

* Convert PDF to HTML (.html)

* Convert PDF to CSV (.csv)

* Convert PDF to Image (.png, .jpg, .jpeg, .jpeg2000, .bmp, .tiff, .tga, .gif, .webp)

* Convert PDF to Plain Text (.txt)

* Convert PDF to Rich Text Format (.rtf)

* Convert PDF to Searchable PDF (.pdf)

* Convert PDF to Structured Data (.json)

* Convert PDF to Markdown (.md)

* Convert PDF to OFD (.ofd)

* Optical Character Recognition (OCR)

* Layout Analysis

* Table Recognition

## Table of Contents

* [Why ComPDF Conversion SDK](#why-compdf-conversion-sdk)

* [Conversion Quality Preview](#conversion-quality-preview)

* [Requirements](#requirements)

* [How to Run a Demo](#how-to-run-a-demo)

* [How to Make a Windows Program in C-sharp with ComPDF Conversion SDK](#how-to-make-a-windows-program-in-c-sharp-with-compdf-conversion-sdk)

* [License and Free Trial](#license-and-free-trial)

* [Developer Guides for All Features](#developer-guides-for-all-features)

* [FAQ](#faq)

* [Changelog](#changelog)

* [Technical Support](#technical-support)

## **Why ComPDF Conversion SDK**

- **Mature Technology:** With years of technology accumulation, we have established a complete mechanism of product iteration to offer a continuous guarantee for product competitiveness.  

- **High-Fidelity Conversion:** Designed to preserve layout, tables, text structure, and formatting in the output files.  

- **AI-Enhanced Conversion:** Continuously improved for complex layout analysis, table recognition, OCR, and editable output. 

- **Independent Intellectual Property Rights:** Our technology is independent and compliant with ISO, helping enterprises conduct international business without considering copyright risks.  

- **High-quality Service:** We provide 24/7 professional one-to-one technical support, including onsite service and remote assistance via phone and email.

## Conversion Quality Preview

With **AI table recognition, AI layout analysis, and OCR capabilities**, ComPDF Conversion SDK delivers accurate PDF conversion results. Simply select the appropriate conversion parameters based on your document type to achieve accurate document reconstruction. Below, we demonstrate the power of the ComPDF Conversion SDK by converting PDFs to Word, Excel, or searchable PDF files. Contact us to get a demo and experience our product's performance firsthand.

![conversion sdk performance GIF](./Images/conversion%20sdk%20performance%20GIF.gif)

### Convert PDF to Word

ComPDF Conversion SDK accurately restores all content and images, complex structures, paragraph aggregation, natural reading order, text styles, bold formatting, fonts, multi‑column layouts, image-text positioning, ordered lists, and more. Get a fully editable Word document that looks exactly like the original PDF.

![conversion sdk performance GIF](./Images/Performance%20-%20PDF%20to%20Word.png)

### Convert PDF to Excel

Extract standard tables and borderless tables with precision, preserve the content and images inside table cells, and retain text styles such as bold and color. ComPDF Conversion SDK turns PDF tables into fully functional Excel spreadsheets without losing structure or formatting.

![Conversion SDK: PDF to Excel](./Images/Performance%20-%20PDF%20to%20Excel.png)

### Convert Scanned PDF to editable text

Convert scanned documents and scanned PDFs into searchable PDFs while perfectly preserving the original layout and formatting. Every word becomes searchable, locatable, and selectable – ready for copying just like native digital text.

![conversion sdk performance GIF](./Images/Performance%20-%20PDF%20to%20searchable%20PDF.png)

## Requirements

| Platform | System Requirements                        | Integrated Development Environment | Note                                                    |
| -------- | ------------------------------------------ | ---------------------------------- | ------------------------------------------------------- |
| Windows  | Windows 7, 8, 10, and 11 (32-bit, 64-bit). | Visual Studio 2017 or higher.      | Framework Requirements: .NET Framework 4.6.1 or higher. |

## How to Run a Demo

ComPDF Conversion SDK for Windows provides a demo in C# for developers to learn how to call the SDK on Windows. You can find the demo in the ***"samples"*** folder. In this guide, it takes the "C#" demo as an example to show how to run it in Visual Studio 2017.

1. Open the demo in Visual Studio 2017: Find and double-click the ***"ComPDFKit_Conversion_Demo.sln"*** in the ***"samples"*** folder.
2. Click **Start** to run the demo on a Windows device. In this guide, a Windows 10 device will be used as an example.

## How to Make a Windows Program in C Sharp with ComPDF Conversion SDK

### Create a New Windows Project

1. Fire up Visual Studio 2017, choose ***File*** -> ***New*** -> ***Project...***, and then select ***Visual C#*** -> ***Windows Desktop*** -> ***Console App(.NET Framework)***.

![conversion sdk](./Images/conversion-windows-make-program-1.png)

2. Choose the options for your new project as picture below. Please make sure to choose .NET Framework 4.6.1 as the programming framework.  

![conversion sdk](./Images/conversion-windows-make-program-2.png)

3. Place the project to the location as desired. Then, click **OK**.

### Add ComPDF Conversion SDK Package

1. Copy all files in the ***"lib"*** folder to the project folder.

2. Add ComPDF Conversion SDK dynamic library to **References**. In order to use ComPDF Conversion SDK APIs in the project, you must add the reference to the project first.
* In Solution Explorer, right-click the project and click **Add -> Reference…**

![conversion sdk](./Images/conversion-windows-add-sdk-1.png)

* In the Add Reference dialog, click the **Browse** tab, navigate to the project folder, select ***"cpdfconversionsdk_dotnet.dll"*** dynamic library, and then click **OK**.

![conversion sdk](./Images/conversion-windows-add-sdk-2.png)

3. Add ComPDF Conversion SDK library to the project. Add the ***"x64"*** and ***"x86"*** folder into the project. Please make sure to set the property **Copy to Output Directory** of ***"cpdfconversionsdk.dll"*** and ***"opencv_world4100.dll"*** to **Copy if newer**. Otherwise, you should copy it to the same folder with the executable file manually before running the project.

![conversion sdk](./Images/conversion-windows-add-sdk-3.png)


## License and Free Trial

### Get Free Trial License

[Contact our sales team](https://www.compdf.com/contact-sales) and we'll send you a 30-day free trial license for ComPDF Conversion SDK.

### Get Commercial License

ComPDF Conversion SDK is a commercial SDK that requires a license for application release. Any documents, sample code, or source code distribution from the released package of ComPDF to any third party is prohibited. To get commercial license for ComPDF Conversion SDK, feel free to [contact our sales team](https://www.compdf.com/contact-sales?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit).

For Windows Conversion SDK, the commercial license must be bound to your developer device ID ([How to find the developer device ID](https://www.compdf.com/faq/how-to-find-the-device-id?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit)), and each license is only valid for one device ID in development mode.

### Apply the License Key

If you haven't get a license key, please check out [how to obtain a license key](/guides/conversion-sdk/csharp/requirements). ComPDF Conversion SDK currently supports offline authentication to verify license keys.

*Learn more:* [*What is the authentication mechanism of ComPDF's license?*](https://www.compdf.com/faq/authentication-mechanism-of-compdfkit-license?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit)

**Copy the License Key:**

Accurately obtaining the license key is crucial for the application of the license.

1. In the email you received, locate the XML file containing the license key.
2. Open the XML file, and determine the license type based on the `<type>` field. If `<type>online</type>` is present, it indicates an online license. If `<type>offline</type>` is present or if the field is absent, it indicates an offline license.
- **Online License:**
  
      <?xml version="1.0" encoding="UTF-8" standalone="no"?>
      <license version="1">
          <platform>windows</platform>
          <starttime>xxxxxxxx</starttime>
          <endtime>xxxxxxxx</endtime>
          <type>online</type>
          <key>LICENSE_KEY</key>
      </license>

- **Offline License**:
  
  ```xml
  <?xml version="1.0" encoding="UTF-8" standalone="no"?>
  <license version="1">
      <platform>linux</platform>
      <starttime>xxxxxxxx</starttime>
      <endtime>xxxxxxxx</endtime>
      <key>LICENSE_KEY</key>
  </license>
  ```
3. Copy the value located at the LICENSE_KEY position within the `<key>LICENSE_KEY</key>` field. This is your license key.
   
   

**Apply the License Key**

You can perform offline authentication using the following method:

Before using the classes and methods of the ComPDF Conversion SDK in your project, you need to initialize the SDK with a valid license key. If you don't have a license key, feel free to contact the [ComPDF team](https://www.compdf.com/contact-sales?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit) to obtain one.

```c#
    string license = " input your license string here ";
    ErrorCode result = LibraryManager.LicenseVerify(license);
```


## Developer Guides for All Features

ComPDF Conversion SDK provides a wide range of customizable conversion options, such as whether to include images or annotations in PDF documents during file conversion, enable OCR, perform layout analysis, and more. Explore the [guides](https://www.compdf.com/guides/conversion-sdk/windows/get-conversion-progress?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit) for sample code and learn how to customize conversions for the following features.

- Get Conversion Progress  

- Cancel Conversion Task  

- Select Page Range  

- Conversion Options: Contain lmage & Annotation  

- Page layout Mode  

- OCR  

- Layout Analysis  

- PDF to Word

- PDF to Excel  

- PDF to PPT  

- PDF to HTML  

- PDF to CSV  

- PDF to lmage  

- PDF to RTF  

- PDF to TXT  

- PDF to searchable PDF  

- Releasing library Resources

### FAQ

* Does OCR work on x86 architecture?
  Currently, the OCR only works on x64 architecture.
- [More FAQ](https://www.compdf.com/faq?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit)

## Changelog

Go to our [changelog](https://www.compdf.com/conversion-sdk/changelog-windows?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit) to keep up with the latest updates, improvements, and bug fixes.

## Technical Support

Thanks for your interest in ComPDF Conversion SDK, the easy-to-use and powerful development solution. If you encounter technical questions or bug issues when using ComPDF Conversion SDK, please submit the problem report to the [ComPDF team](mailto:support@compdf.com). More information as follows would help us to solve your problem:

- ComPDF Conversion SDK product and version.
- Your operating system and IDE version.
- Detailed descriptions of the problem.
- Any other related information, such as an error screenshot.
  
  

Home link: [https://www.compdf.com](https://www.compdf.com/?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit)

Technical Support: [https://www.compdf.com/support](https://www.compdf.com/support?utm_source=github&utm_medium=compdfkit-conversion-sdk-windows&utm_campaign=compdfkit_conversion_sdk_windows_repo&ref_platform_id=github_compdfkit)

Email: [support@compdf.com](mailto:support@compdf.com)

