package de.fricofresh.outlookspitter;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class StartGUI extends Application {
	
	@Override
	public void start(Stage stage) throws Exception {
		
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("gui/MainWindow.fxml"));
		// MainWindowController controller = fxmlLoader.getController();
		stage.setScene(new Scene(fxmlLoader.load()));
		stage.show();
	}
	
	public static void main(String[] args) {
		
		launch(args);
	}
}
