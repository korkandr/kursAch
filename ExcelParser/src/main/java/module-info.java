module com.excelparser {
    requires javafx.controls;
    requires javafx.fxml;

    requires org.controlsfx.controls;
    requires org.kordamp.bootstrapfx.core;
    requires org.apache.logging.log4j;
    requires aspose.cells;

    opens com.excelparser to javafx.fxml;
    exports com.excelparser;
}