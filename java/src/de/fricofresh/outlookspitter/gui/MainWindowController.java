package de.fricofresh.outlookspitter.gui;

/**
 * Sample Skeleton for 'MainWindow.fxml' Controller Class
 */

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.ResourceBundle;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hsmf.MAPIMessage;

import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;
import de.fricofresh.outlookspitter.CreateSplittedFilesParameter;
import de.fricofresh.outlookspitter.MailGenMethod;
import de.fricofresh.outlookspitter.utils.JavaMailMessageUtil;
import de.fricofresh.outlookspitter.utils.MailSplitterUtil;
import de.fricofresh.outlookspitter.utils.OutlookSplitterProcessorUtil;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.TextField;

public class MainWindowController {
	
	@FXML // ResourceBundle that was given to the FXMLLoader
	private ResourceBundle resources;
	
	@FXML // URL location of the FXML file that was given to the FXMLLoader
	private URL location;
	
	@FXML // fx:id="splitTextField"
	private TextField splitTextField; // Value injected by FXMLLoader
	
	@FXML // fx:id="recipientsFilePickerButton"
	private Button recipientsToFilePickerButton; // Value injected by FXMLLoader
	
	@FXML // fx:id="recipientsTextfield"
	private TextField recipientsToTextfield; // Value injected by FXMLLoader
	
	@FXML // fx:id="recipientsFilePickerButton"
	private Button recipientsCCFilePickerButton; // Value injected by FXMLLoader
	
	@FXML // fx:id="recipientsTextfield"
	private TextField recipientsCCTextfield; // Value injected by FXMLLoader
	
	@FXML // fx:id="recipientsFilePickerButton"
	private Button recipientsBCCFilePickerButton; // Value injected by FXMLLoader
	
	@FXML // fx:id="recipientsTextfield"
	private TextField recipientsBCCTextfield; // Value injected by FXMLLoader
	
	@FXML // fx:id="baseMailFilePickerButton"
	private Button baseMailFilePickerButton; // Value injected by FXMLLoader
	
	@FXML // fx:id="baseMailTextField"
	private TextField baseMailTextField; // Value injected by FXMLLoader
	
	@FXML // fx:id="emailHTMLFilePickerButton"
	private Button emailHTMLFilePickerButton; // Value injected by FXMLLoader
	
	@FXML // fx:id="emailHTMLTextField"
	private TextField emailHTMLTextField; // Value injected by FXMLLoader
	
	@FXML // fx:id="outputFilePickerButton"
	private Button outputFilePickerButton; // Value injected by FXMLLoader
	
	@FXML // fx:id="outputTextField"
	private TextField outputTextField; // Value injected by FXMLLoader
	
	@FXML // fx:id="outlookExeTextField"
	private TextField outlookExeTextField; // Value injected by FXMLLoader
	
	@FXML // fx:id="outlookExeFilePickerButton"
	private Button outlookExeFilePickerButton; // Value injected by FXMLLoader
	
	@FXML // fx:id="openAfterCreationCheckBox"
	private CheckBox openAfterCreationCheckBox; // Value injected by FXMLLoader
	
	@FXML // fx:id="createButton"
	private Button createButton; // Value injected by FXMLLoader
	
	private Logger log = LogManager.getLogger(MainWindowController.class);
	
	@FXML
	void createOutlookFiles(ActionEvent event) {
		
		try {
			CreateSplittedFilesParameter csfParameter = new CreateSplittedFilesParameter();
			
			String recipientsToText = recipientsToTextfield.getText();
			String recipientsCCText = recipientsCCTextfield.getText();
			String recipientsBCCText = recipientsBCCTextfield.getText();
			List<OutlookMessageRecipient> emails = new ArrayList<>();
			if (new File(recipientsToText).exists()) {
				emails.addAll(MailSplitterUtil.getOutlookRecipientsList(MailSplitterUtil.extractMailsFromFile(new File(recipientsToText)), Type.TO));
			}
			else if (recipientsToText.contains(";")) {
				emails.addAll(OutlookSplitterProcessorUtil.receiveOutlookRecipients(recipientsToText, Type.TO));
			}
			if (new File(recipientsCCText).exists()) {
				emails.addAll(MailSplitterUtil.getOutlookRecipientsList(MailSplitterUtil.extractMailsFromFile(new File(recipientsToText)), Type.CC));
			}
			else if (recipientsCCText.contains(";")) {
				emails.addAll(OutlookSplitterProcessorUtil.receiveOutlookRecipients(recipientsCCText, Type.CC));
			}
			if (new File(recipientsBCCText).exists()) {
				emails.addAll(MailSplitterUtil.getOutlookRecipientsList(MailSplitterUtil.extractMailsFromFile(new File(recipientsToText)), Type.BCC));
			}
			else if (recipientsBCCText.contains(";")) {
				emails.addAll(OutlookSplitterProcessorUtil.receiveOutlookRecipients(recipientsBCCText, Type.BCC));
			}
			
			if (emails.isEmpty()) {
				new AdvancedAlert("Fehler beim der Liste der E-Mails", "Bitte gebe eine Datei oder eine Liste, welche die E-Mails aus Outlook herauskopiert wurden an.", AlertType.ERROR).show();
				return;
			}
			
			csfParameter.setRecipientsToSplit(emails);
			try {
				csfParameter.setSplit(Integer.valueOf(splitTextField.getText()));
			}
			catch (Exception e) {
				new AdvancedAlert("Fehler bitte geben Sie eine Anzahl an Empfänger an", "", AlertType.ERROR);
				return;
			}
			
			csfParameter.setEmailPath(Path.of(baseMailTextField.getText()));
			try {
				csfParameter.setEmailMessage(new MAPIMessage(new File(baseMailTextField.getText())));
			}
			catch (IOException e) {
				log.error(e, e);
			}
			csfParameter.setEmailHTMLMessage(Optional.ofNullable(emailHTMLTextField.getText()));
			csfParameter.setOutputDir(Optional.ofNullable(outputTextField.getText()));
			csfParameter.setMailGenMehtod(MailGenMethod.JAVAMAIL);
			List<Path> splittedFiles = JavaMailMessageUtil.createSplittedFiles(csfParameter);
			
			if (openAfterCreationCheckBox.isSelected()) {
				MailSplitterUtil.openFiles(splittedFiles, outlookExeTextField.getText().isBlank() ? Optional.empty() : Optional.ofNullable(outlookExeTextField.getText()));
			}
		}
		catch (FileAlreadyExistsException e) {
			log.error(e, e);
			new AdvancedAlert("Ein Fehler ist aufgetreten", "Dateien Existiert schon", "Die unter 'Dateien erstellt in' angegebenen Ordner enthält bereits generierte E-Mails, bitte lösche diese und generieren Sie erneut.",
					AlertType.ERROR);
			return;
			
		}
		catch (Exception e) {
			log.error(e, e);
			new AdvancedAlert("Ein Fehler ist aufgetreten", e.toString(), AlertType.ERROR);
			return;
		}
		new AdvancedAlert("Erfolgreich", "Die Dateien wurden erstellt", AlertType.INFORMATION);
	}
	
	@FXML
	void handleBaseMailFilePicker(ActionEvent event) {
		
	}
	
	@FXML
	void handleEmailHTMLFilePicker(ActionEvent event) {
		
	}
	
	@FXML
	void handleOutlookExeFilePicker(ActionEvent event) {
		
	}
	
	@FXML
	void handleOutputFileFilePicker(ActionEvent event) {
		
	}
	
	@FXML
	void handleToRecipientsFilePicker(ActionEvent event) {
		
	}
	
	@FXML
	void handleCCRecipientsFilePicker(ActionEvent event) {
		
	}
	
	@FXML
	void handleBCCRecipientsFilePicker(ActionEvent event) {
		
	}
	
	@FXML // This method is called by the FXMLLoader when initialization is complete
	void initialize() {
		
		assert splitTextField != null : "fx:id=\"splitTextField\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert recipientsToFilePickerButton != null : "fx:id=\"recipientsFilePickerButton\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert recipientsToTextfield != null : "fx:id=\"recipientsTextfield\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert recipientsCCFilePickerButton != null : "fx:id=\"recipientsFilePickerButton\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert recipientsCCTextfield != null : "fx:id=\"recipientsTextfield\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert recipientsBCCFilePickerButton != null : "fx:id=\"recipientsFilePickerButton\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert recipientsBCCTextfield != null : "fx:id=\"recipientsTextfield\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert baseMailFilePickerButton != null : "fx:id=\"baseMailFilePickerButton\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert baseMailTextField != null : "fx:id=\"baseMailTextField\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert emailHTMLFilePickerButton != null : "fx:id=\"emailHTMLFilePickerButton\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert emailHTMLTextField != null : "fx:id=\"emailHTMLTextField\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert outputFilePickerButton != null : "fx:id=\"outputFilePickerButton\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert outputTextField != null : "fx:id=\"outputTextField\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert outlookExeTextField != null : "fx:id=\"outlookExeTextField\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert outlookExeFilePickerButton != null : "fx:id=\"outlookExeFilePickerButton\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert openAfterCreationCheckBox != null : "fx:id=\"openAfterCreationCheckBox\" was not injected: check your FXML file 'MainWindow.fxml'.";
		assert createButton != null : "fx:id=\"createButton\" was not injected: check your FXML file 'MainWindow.fxml'.";
		
	}
}
