package com.excelparser;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

/**
 * Класс представляет собой точку входа для приложения ExcelParser.
 * Он расширяет класс приложения JavaFX и определяет метод start
 * для инициализации и отображения основного пользовательского интерфейса приложения.
 */
public class Main extends Application {

    /**
     * Метод start вызывается при запуске приложения JavaFX.
     * Он загружает файл FXML из ресурсов, создает сцену и устанавливает
     * основное окно приложения с заданными размерами. Затем отображает основное окно.
     *
     * @param stage основное окно приложения
     * @throws IOException если произошла ошибка при загрузке файла FXML.
     */

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(Main.class.getResource("main.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 495, 355);
        stage.setTitle("ExcelParser");
        stage.setScene(scene);
        stage.show();
    }

    /**
     * Метод main является точкой входа для Java-приложения.
     * Он запускает приложение JavaFX, вызывая метод {@code launch()}.
     *
     * @param args аргументы командной строки (не используются в данном приложении).
     */

    public static void main(String[] args) {
        launch();
    }
}