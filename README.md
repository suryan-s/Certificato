# Certificato

Certificato is a web app developed using Streamlit to help users send certificates in large quantities when provided with an excel sheet containing the name and email address of the recipients.

## 📝 Features

* The app reads the Excel sheet, extracts the name of the recipient, and creates a customized certificate with the provided name. The certificate is then sent to the recipient's email address.

* The app is designed to use the app password of the Google account instead of the usual account password, as per the new Google security rule which doesn't allow third-party integration. This ensures that the user's account is protected and secure.

* In addition to the current functionality, the future updates of Certificato will include the option to use sample certificates and templates and make edits on those. This feature will provide users with more flexibility in creating certificates and personalizing them according to their needs.

* One of the key features of Certificato is its simplicity and ease of use. Users can upload their Excel sheet, select the certificate template, and send the certificates with just a few clicks. The app takes care of the rest, including creating the certificates and sending them to the recipient's email address.

* Another advantage of Certificato is its compatibility with both Windows and Linux servers. This makes it a versatile app that can be used in various environments and by users with different technical backgrounds.

## 🚀 Getting Started

* Clone this repository by running the following command:

    `git clone https://github.com/<username>/certificato.git`

* Install the required packages using the following command:

    `pip install -r requirements.txt`
* Run the following command to start the web app:

    `streamlit run server.py`

## 👩‍💻 Usage

* Once the app is running, upload the Excel sheet containing the name and email address of recipients.
* Select the certificate template from local device or use sample certificate.
* The certificate could be created by following this procedure:
    1. Create a certificate and export it as PDF.
    2. Open the certificate in Word and create a text box within the certificate where the name of the participant is to be created.
    3. Within the text box add the jinja syntax : **{{value}}**
    4. It's important to note that as far as the current version of the project, the application is able to edit / add only the name of the participant. the rest of the body for the certificate have to be finalised before converting to PDF.
    5. Make sure the textbox width touches both the extreme left and right side as in Ubuntu, the covertion is done by libreoffice and it was found that miss alignment is common if the width is not set to either side. An example is as shown below:
    ![Screenshot 2023-03-22 142041](https://user-images.githubusercontent.com/76394506/226849663-c88463e8-cb99-4e5f-9a33-a30303a7e76a.png)

* (*Imp*) A sample certificate to be fed into the program would be present in the root as cert.docx . The Jinja syntax within the certificate would be replaced with the name of the participant mentioned in the excel sheet. Also when the mail is sent, if the body of the mail have '#' in it, then it would be replaced with the name of the participant.
* (*Imp*) If the certificate is made in Canva, make sure you export the pdf as 'PDF Print' so that you won't face any issue while opening the file in word to add the Jinja syntax.
* Click on the 'Send Certificates' button.
* Certificates will be generated and sent to the respective email addresses.
* A .zip file would be available to download the whole certificates generated and an excel sheel would be also be there which shows the participant name and the status if the mail was successfully sent or not.

## 👨‍💻 Contributing

* Certificato is an open-source project, which means that users can contribute to the development of the app and suggest new features and improvements. This also ensures that the app is continuously updated and maintained by the community.
* Contributions are always welcome! Please create a pull request with your changes.

## 📝 License

This project is licensed under the MIT License.

## 📧 Contact

If you have any questions or suggestions, please feel free to contact me at suryannasa@gmail.com
