/*
 * Copyright (C) 2014 captainbowtie
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package com.AllenBarr.CallSheetGenerator;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;
import java.net.URL;
import java.util.Date;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.layout.BorderPane;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.supercsv.cellprocessor.ParseDate;
import org.supercsv.cellprocessor.constraint.NotNull;
import org.supercsv.cellprocessor.ift.CellProcessor;
import org.supercsv.io.CsvListReader;
import org.supercsv.io.ICsvListReader;
import org.supercsv.prefs.CsvPreference;

/**
 *
 * @author captainbowtie
 */
public class CallSheetGeneratorControllerController implements Initializable {

    final private File configFile;
    private File excelSheet;
    private File contributionDirectory;
    private File callSheetDirectory;
    private Workbook wb;
    private Sheet wbSheet;
    final Pattern dateRegex = Pattern.compile("([A-z].. [A-z].. [0-9]. [0-9:]....... [A-Z].. [0-9]...)");
    final Pattern descRegex = Pattern.compile("(\\#[0-9 [A-z]]*)");
    final Pattern amtRegex = Pattern.compile("([0-9]*\\.[0-9].)");
    @FXML
    private BorderPane rootPane;
    @FXML
    private Button generateButton;
    @FXML
    private Button generateAtButton;
    @FXML
    private Button locateExcelButton;
    @FXML
    private Button locateDirectoryButton;
    @FXML
    private Label excelLabel;
    @FXML
    private Label directoryLabel;
    @FXML
    private ComboBox contributorSelector;

    public CallSheetGeneratorControllerController() {
        final Thread mainThread = Thread.currentThread();
        Runtime.getRuntime().addShutdownHook(new Thread() {
            public void run() {
                saveConfig();
            }
        });
        configFile = new File("CSG.cfg");

    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        
        generateAtButton.setDisable(true);
        generateButton.setDisable(true);
        if (configFile.exists()) {
            try {
                try (final Scanner config = new Scanner(configFile)) {
                    String line;
                    if (!"noExcel".equals(line = config.nextLine())) {
                        excelSheet = new File(line);
                        excelLabel.setText(excelSheet.getName());
                        populateContributorList();
                    }
                    if (!"noFolder".equals(line = config.nextLine())) {
                        contributionDirectory = new File(line);
                        directoryLabel.setText(contributionDirectory.getName());
                    }
                    if (!"noSaveFolder".equals(line = config.nextLine())) {
                        callSheetDirectory = new File(line);
                    }
                    config.close();
                }
            } catch (FileNotFoundException e) {

            }
        }

    }

    @FXML
    private void handleGenerateButton(ActionEvent e) {
        if (contributorSelector.getSelectionModel().getSelectedIndex() != 0) {
            final File csFile = new File(callSheetDirectory.toString() + "/" + ((String) contributorSelector.getItems().get(contributorSelector.getSelectionModel().getSelectedIndex())).replaceAll("[\\D]", "") + ".xlsx");
            final Generator csGen = new Generator(csFile, makeContributor());
        }
    }

    @FXML
    private void handleGenerateAtButton(ActionEvent e) {
        if (contributorSelector.getSelectionModel().getSelectedIndex() != 0) {
            final FileChooser fc = new FileChooser();
            fc.setTitle("Save Call Sheet...");
            fc.getExtensionFilters().addAll(new ExtensionFilter("Excel Workbook (.xlsx)", "*.xlsx"), new ExtensionFilter("Excel 97-2007 Workbook (.xls)", "*.xls"));
            final File fcFile = fc.showSaveDialog(new Stage());
            if (fcFile != null && !fcFile.exists()) {
                callSheetDirectory = fcFile.getParentFile();
                generateButton.setDisable(false);

                final Generator csGen = new Generator(fcFile, makeContributor());
            }
        }
    }

    @FXML
    private void handleFindExcelButton(ActionEvent e) {
        final FileChooser fc = new FileChooser();
        fc.setTitle("Please locate contributor database...");
        final File fcFile = fc.showOpenDialog(new Stage());
        if (null != fcFile) {
            if (fcFile.exists()) {
                excelSheet = fcFile;
                excelLabel.setText(excelSheet.getName());
                populateContributorList();
            }
        } else {
        }
    }

    @FXML
    private void handleFindDirectoryButton(ActionEvent e) {
        final DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle("Please locate the contribution record directory...");
        final File dcDirectory = dc.showDialog(new Stage());
        if (null != dcDirectory) {
            if (dcDirectory.exists()) {
                contributionDirectory = dcDirectory;
                directoryLabel.setText(contributionDirectory.getName());
            }
        }
    }

    @FXML
    private void handleQuitMenu(ActionEvent e) {
        shutdown();
    }

    @FXML
    private void handleAboutMenu(ActionEvent e) {
        Parent root = null;
        try {
            root = FXMLLoader.load(getClass().getResource("AboutWindowFXML.fxml"));
        } catch (IOException ex) {
            Logger.getLogger(CallSheetGeneratorControllerController.class.getName()).log(Level.SEVERE, null, ex);
        }
        final Scene scene;
        scene = new Scene(root);
        final Stage aboutStage = new Stage();
        aboutStage.setScene(scene);
        aboutStage.show();
    }

    @FXML
    private void handleChooserSelection(ActionEvent e){
        if(contributorSelector.getSelectionModel().getSelectedIndex()!=0){
            generateAtButton.setDisable(false);
            if(callSheetDirectory!=null && callSheetDirectory.exists()){
                generateButton.setDisable(false);
            }
        }else{
            generateButton.setDisable(true);
            generateAtButton.setDisable(true);
        }
    }
    
    private void shutdown() {
        saveConfig();
        System.exit(0);
    }

    private void saveConfig() {
        PrintWriter config = null;
        try {
            config = new PrintWriter(configFile);

            if (null != excelSheet && excelSheet.exists()) {
                config.println(excelSheet.toString());
            } else {
                config.println("noExcel");
            }
            if (null != contributionDirectory && contributionDirectory.exists()) {
                config.println(contributionDirectory.toString());
            } else {
                config.println("noFolder");
            }
            if (null != callSheetDirectory && callSheetDirectory.exists()) {
                config.println(callSheetDirectory.toString());
            } else {
                config.println("noSaveFolder");
            }
            config.close();

        } catch (Exception e) {
            e.printStackTrace();
            config.close();
        }
    }

    private void populateContributorList() {
        if (excelSheet.exists()) {
            try {
                wb = WorkbookFactory.create(excelSheet);
                wbSheet = wb.getSheetAt(0);
                Row wbRow = wbSheet.getRow(0);
                Integer vanIDColumnIndex = 0;
                Integer fNameColumnIndex = 0;
                Integer lNameColumnIndex = 0;
                for (Cell cell : wbRow) {
                    if (null != cell.getStringCellValue()) {
                        switch (cell.getStringCellValue()) {
                            case "VANID":
                                vanIDColumnIndex = cell.getColumnIndex();
                                break;
                            case "LastName":
                                lNameColumnIndex = cell.getColumnIndex();
                                break;
                            case "FirstName":
                                fNameColumnIndex = cell.getColumnIndex();
                                break;
                        }
                    }
                }
                final ObservableList<String> names = FXCollections.observableArrayList();
                for (Row row : wbSheet) {
                    switch (row.getCell(vanIDColumnIndex).getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            names.add(row.getCell(vanIDColumnIndex).getStringCellValue() + " " + row.getCell(fNameColumnIndex).getStringCellValue() + " " + row.getCell(lNameColumnIndex).getStringCellValue());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            names.add((int) row.getCell(vanIDColumnIndex).getNumericCellValue() + " " + row.getCell(fNameColumnIndex).getStringCellValue() + " " + row.getCell(lNameColumnIndex).getStringCellValue());
                            break;
                    }
                }
                contributorSelector.setItems(names);
                contributorSelector.getSelectionModel().select(0);
            } catch (IOException | InvalidFormatException ex) {
                Logger.getLogger(CallSheetGeneratorControllerController.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

    private Contributor makeContributor() {
        Contributor contributor = new Contributor();
        Row contribRow = wbSheet.getRow(contributorSelector.getSelectionModel().getSelectedIndex());
        Row headerRow = wbSheet.getRow(0);
        Integer vanIDColumnIndex = 0;
        Integer lNameColumnIndex = 0;
        Integer fNameColumnIndex = 0;
        Integer mNameColumnIndex = 0;
        Integer suffixColumnIndex = 0;
        Integer salutationColumnIndex = 0;
        Integer spouseColumnIndex = 0;
        Integer mAddressColumnIndex = 0;
        Integer mCityColumnIndex = 0;
        Integer mStateColumnIndex = 0;
        Integer mZipColumnIndex = 0;
        Integer sexColumnIndex = 0;
        Integer ageColumnIndex = 0;
        Integer occupationColumnIndex = 0;
        Integer employerColumnIndex = 0;
        Integer phoneColumnIndex = 0;
        Integer homePhoneColumnIndex = 0;
        Integer cellPhoneColumnIndex = 0;
        Integer workPhoneColumnIndex = 0;
        Integer extColumnIndex = 0;
        Integer emailColumnIndex = 0;
        Integer notesColumnIndex = 0;
        Integer partyColumnIndex = 0;
        for (Cell cell : headerRow) {
            if (null != cell.getStringCellValue()) {
                switch (cell.getStringCellValue()) {
                    case "VANID":
                        vanIDColumnIndex = cell.getColumnIndex();
                        break;
                    case "LastName":
                        lNameColumnIndex = cell.getColumnIndex();
                        break;
                    case "FirstName":
                        fNameColumnIndex = cell.getColumnIndex();
                        break;
                    case "MiddleName":
                        mNameColumnIndex = cell.getColumnIndex();
                        break;
                    case "Suffix":
                        suffixColumnIndex = cell.getColumnIndex();
                        break;
                    case "Salutation":
                        salutationColumnIndex = cell.getColumnIndex();
                        break;
                    case "Spouse":
                        spouseColumnIndex = cell.getColumnIndex();
                        break;
                    case "mAddress":
                        mAddressColumnIndex = cell.getColumnIndex();
                        break;
                    case "mCity":
                        mCityColumnIndex = cell.getColumnIndex();
                        break;
                    case "mState":
                        mStateColumnIndex = cell.getColumnIndex();
                        break;
                    case "mZip5":
                        mZipColumnIndex = cell.getColumnIndex();
                        break;
                    case "Sex":
                        sexColumnIndex = cell.getColumnIndex();
                        break;
                    case "Age":
                        ageColumnIndex = cell.getColumnIndex();
                        break;
                    case "Occupation":
                        occupationColumnIndex = cell.getColumnIndex();
                        break;
                    case "Employer":
                        employerColumnIndex = cell.getColumnIndex();
                        break;
                    case "Phone":
                        phoneColumnIndex = cell.getColumnIndex();
                        break;
                    case "CellPhone":
                        cellPhoneColumnIndex = cell.getColumnIndex();
                        break;
                    case "HomePhone":
                        homePhoneColumnIndex = cell.getColumnIndex();
                        break;
                    case "WorkPhone":
                        workPhoneColumnIndex = cell.getColumnIndex();
                        break;
                    case "WorkPhoneExt":
                        extColumnIndex = cell.getColumnIndex();
                        break;
                    case "Email":
                        emailColumnIndex = cell.getColumnIndex();
                        break;
                    case "Notes":
                        notesColumnIndex = cell.getColumnIndex();
                        break;
                    case "Party":
                        partyColumnIndex = cell.getColumnIndex();
                        break;
                }
            }
        }
        try {
            contributor.setVANID((int) contribRow.getCell(vanIDColumnIndex).getNumericCellValue());
        } catch (NullPointerException ex) {

        } catch (IllegalStateException ex) {
            contributor.setVANID(Integer.parseInt(contribRow.getCell(vanIDColumnIndex).getStringCellValue()));
        }
        String name = "";
        try {
            name = name.concat(contribRow.getCell(fNameColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            name = name.concat(" " + contribRow.getCell(mNameColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            name = name.concat(" " + contribRow.getCell(lNameColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            name = name.concat(" " + contribRow.getCell(suffixColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        contributor.setName(name);
        try {
            contributor.setSalutation(contribRow.getCell(salutationColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setSex(contribRow.getCell(sexColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setParty(contribRow.getCell(partyColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setPhone(contribRow.getCell(phoneColumnIndex).getNumericCellValue());
        } catch (NullPointerException ex) {

        } catch (IllegalStateException ex) {
            contributor.setPhone(Double.parseDouble(contribRow.getCell(phoneColumnIndex).getStringCellValue()));
        }
        try {
            contributor.setHomePhone(contribRow.getCell(homePhoneColumnIndex).getNumericCellValue());
        } catch (NullPointerException ex) {

        } catch (IllegalStateException ex) {
            contributor.setHomePhone(Double.parseDouble(contribRow.getCell(homePhoneColumnIndex).getStringCellValue()));
        }
        try {
            contributor.setCellPhone(contribRow.getCell(cellPhoneColumnIndex).getNumericCellValue());
        } catch (NullPointerException ex) {

        } catch (IllegalStateException ex) {
            contributor.setCellPhone(Double.parseDouble(contribRow.getCell(cellPhoneColumnIndex).getStringCellValue()));
        }
        try {
            contributor.setWorkPhone(contribRow.getCell(workPhoneColumnIndex).getNumericCellValue());
        } catch (NullPointerException ex) {

        } catch (IllegalStateException ex) {
            contributor.setWorkPhone(Double.parseDouble(contribRow.getCell(workPhoneColumnIndex).getStringCellValue()));
        }
        try {
            contributor.setWorkExtension(contribRow.getCell(extColumnIndex).getNumericCellValue());
        } catch (NullPointerException ex) {

        } catch (IllegalStateException ex) {
            contributor.setWorkExtension(Double.parseDouble(contribRow.getCell(extColumnIndex).getStringCellValue()));
        }
        try {
            contributor.setEmail(contribRow.getCell(emailColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setEmployer(contribRow.getCell(employerColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setOccupation(contribRow.getCell(occupationColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setAge((int) contribRow.getCell(ageColumnIndex).getNumericCellValue());
        } catch (NullPointerException ex) {

        } catch (IllegalStateException ex) {
            contributor.setAge(Integer.parseInt(contribRow.getCell(ageColumnIndex).getStringCellValue()));
        }
        try {
            contributor.setSpouse(contribRow.getCell(spouseColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setStreetAddress(contribRow.getCell(mAddressColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setCity(contribRow.getCell(mCityColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setState(contribRow.getCell(mStateColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        try {
            contributor.setZip((int) contribRow.getCell(mZipColumnIndex).getNumericCellValue());
        } catch (NullPointerException ex) {

        } catch (IllegalStateException ex) {
            contributor.setZip(Integer.parseInt(contribRow.getCell(mZipColumnIndex).getStringCellValue()));
        }
        try {
            contributor.setNotes(contribRow.getCell(notesColumnIndex).getStringCellValue());
        } catch (NullPointerException ex) {

        }
        ICsvListReader listReader = null;
        try {
            try {
                listReader = new CsvListReader(new FileReader(contributionDirectory.getPath() + "/" + ((String) contributorSelector.getItems().get(contributorSelector.getSelectionModel().getSelectedIndex())).replaceAll("[\\D]", "") + ".csv"), CsvPreference.STANDARD_PREFERENCE);
            listReader.getHeader(true); // skip the header (can't be used with CsvListReader)

            while ((listReader.read()) != null) {

                // use different processors depending on the number of columns
                final CellProcessor[] processors;
                processors = getProcessors();

                final List<Object> customerList = listReader.executeProcessors(processors);

                Matcher date = dateRegex.matcher(customerList.toString());
                Matcher des = descRegex.matcher(customerList.toString());
                Matcher amt = amtRegex.matcher(customerList.toString());
                date.find();
                des.find();
                amt.find();
                contributor.addDonation(new Donation(new Date(date.group(1)), des.group(1), Double.parseDouble(amt.group(1))));
            }
            
            } catch (FileNotFoundException | NullPointerException ex) {
                //Logger.getLogger(CallSheetGeneratorControllerController.class.getName()).log(Level.SEVERE, null, ex);
            }

            

        } catch (IOException ex) {
            //Logger.getLogger(CallSheetGeneratorControllerController.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            if (listReader != null) {
                try {
                    listReader.close();
                } catch (IOException ex) {
                    //Logger.getLogger(CallSheetGeneratorControllerController.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }
        return contributor;
    }

    private static CellProcessor[] getProcessors() {

        final CellProcessor[] processors = new CellProcessor[]{
            new NotNull(), // type
            new ParseDate("MM/dd/yy"), // date
            new NotNull(), // name
            new NotNull(), // street address
            new NotNull(), // city
            new NotNull(), // State
            new NotNull(), // zip
            new NotNull(), // recipient
            new NotNull(), //recip address
            new NotNull(), // recip city
            new NotNull(), // recip state
            new NotNull(), // recip zip
            new NotNull(), // amount
        };

        return processors;
    }
}
