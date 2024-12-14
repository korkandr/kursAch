package com.excelparser;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;

/**
 * Класс представляет контроллер для графического интерфейса приложения ExcelParser.
 * Он содержит аннотации FXML для связывания элементов интерфейса с соответствующими полями и методами контроллера.
 */
public class MainController {

    /**
     * Кнопка, обрабатывающая событие нажатия.
     */

    @FXML
    private Button button;

    /**
     * Текстовое поле для ввода данных.
     */

    @FXML
    private TextField field_input;

    /**
     * Текстовое поле для ввода данных.
     */

    @FXML
    private TextField field_output;

    /**
     * Метод вызывается при инициализации контроллера.
     * Он устанавливает обработчик события для кнопки, реагирующий на нажатие.
     */

    @FXML
    void initialize() {
        button.setOnAction(actionEvent -> {
            if (!field_input.getText().isEmpty() && !field_output.getText().isEmpty()) {
                ExcelParser.start(field_input.getText(), field_output.getText());
                field_input.setText("");
                field_output.setText("");
            }
        });
    }
}