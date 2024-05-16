package de.fricofresh.outlookspitter.gui;

import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.stage.Stage;

public class AdvancedAlert extends Alert {
	
	/**
	 * Öffnet ein Alertfenster.
	 * 
	 * @param headerText
	 * @param content
	 * @param type
	 */
	public AdvancedAlert(String headerText, String content, AlertType alertType) {
		
		super(alertType);
		if (AlertType.CONFIRMATION.equals(alertType))
			setTitle("Bestätigung");
		else if (AlertType.ERROR.equals(alertType))
			setTitle("Fehler");
		else if (AlertType.INFORMATION.equals(alertType))
			setTitle("Information");
		else if (AlertType.WARNING.equals(alertType))
			setTitle("Warnung");
		else
			setTitle("Meldung");
		setHeaderText(headerText);
		setContentText(content);
		showAndWait();
	}
	
	/**
	 * �ffnet ein Alertfenster.
	 * 
	 * @param headerText
	 * @param content
	 * @param type
	 */
	public AdvancedAlert(String title, String headerText, String content, AlertType alertType) {
		
		super(alertType);
		setTitle(title);
		setHeaderText(headerText);
		setContentText(content);
		showAndWait();
	}
	
	/**
	 * Schlie�t ein Alertfenster ohne Buttons (Typ "loading").
	 * 
	 * @author Flo
	 */
	public void closeAlert() {
		
		getButtonTypes().add(ButtonType.CANCEL);
		hide();
		getButtonTypes().remove(ButtonType.CANCEL);
	}
	
	public void addButtonTypes(ButtonType... buttonTypes) {
		
		getButtonTypes().addAll(buttonTypes);
	}
	
	public void setButtonTypes(ButtonType... buttonTypes) {
		
		getButtonTypes().setAll(buttonTypes);
	}
	
	public void setFocus(boolean focus) {
		
		((Stage) getDialogPane().getScene().getWindow()).setAlwaysOnTop(focus);
	}
}
