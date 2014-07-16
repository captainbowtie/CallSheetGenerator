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
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author captainbowtie
 */
public class Generator {

    public Generator(File file, Contributor contrib) {
        generateSheet(file, contrib);
    }

    public int generateSheet(File file, Contributor contrib) {
        //create workbook file
        final String fileName = file.toString();
        final Workbook wb;
        if (fileName.endsWith(".xlsx")) {
            wb = new XSSFWorkbook();
        } else if (fileName.endsWith(".xls")) {
            wb = new HSSFWorkbook();
        } else {
            return 1;
        }
        //create sheet
        final Sheet sheet = wb.createSheet("Call Sheet");
        final Header header = sheet.getHeader();
        header.setCenter("Anderson for Iowa Call Sheet");
        //add empty cells
        final Row[] row = new Row[22 + contrib.getDonationsLength()];
        final Cell[][] cell = new Cell[6][22 + contrib.getDonationsLength()];
        for (int i = 0; i < (22 + contrib.getDonationsLength()); i++) {
            row[i] = sheet.createRow((short) i);
            for (int j = 0; j < 6; j++) {
                cell[j][i] = row[i].createCell(j);
            }
        }
        //populate cells with data
        //column 1
        cell[0][0].setCellValue(contrib.getName());
        cell[0][3].setCellValue("Sex:");
        cell[0][4].setCellValue("Party:");
        cell[0][5].setCellValue("Phone #:");
        cell[0][6].setCellValue("Home #:");
        cell[0][7].setCellValue("Cell #:");
        cell[0][8].setCellValue("Work #:");
        cell[0][10].setCellValue("Email:");
        cell[0][12].setCellValue("Employer:");
        cell[0][13].setCellValue("Occupation:");
        cell[0][15].setCellValue("Past Contact:");
        cell[0][17].setCellValue("Notes:");
        cell[0][21].setCellValue("Contribution History:");
        //column 2
        cell[1][3].setCellValue(contrib.getSex());
        cell[1][4].setCellValue(contrib.getParty());
        cell[1][5].setCellValue(contrib.getPhone());
        cell[1][6].setCellValue(contrib.getHomePhone());
        cell[1][7].setCellValue(contrib.getCellPhone());
        cell[1][8].setCellValue(contrib.getWorkPhone());
        cell[1][9].setCellValue("x" + contrib.getWorkExtension());
        cell[1][10].setCellValue(contrib.getEmail());
        cell[1][12].setCellValue(contrib.getEmployer());
        cell[1][13].setCellValue(contrib.getOccupation());
        cell[1][17].setCellValue(contrib.getNotes());
        //column 4
        cell[3][3].setCellValue("Salutation:");
        cell[3][4].setCellValue("Age:");
        cell[3][5].setCellValue("Spouse:");
        cell[3][7].setCellValue("Address:");
        cell[3][10].setCellValue("TARGET:");
        //column 5
        cell[4][0].setCellValue("VANID:");
        cell[4][3].setCellValue(contrib.getSalutation());
        cell[4][4].setCellValue(contrib.getAge());
        cell[4][5].setCellValue(contrib.getSpouse());
        cell[4][7].setCellValue(contrib.getStreetAddress());
        cell[4][8].setCellValue(contrib.getCity() + ", " + contrib.getState() + " " + contrib.getZip());
        //column 6
        cell[5][0].setCellValue(contrib.getVANID());
        //contribution cells
        for (int i = 0; i < contrib.getDonationsLength(); i++) {
            cell[0][i + 22].setCellValue(contrib.getDonation(i).getDonationDate());
            cell[1][i + 22].setCellValue(contrib.getDonation(i).getRecipient());
            cell[5][i + 22].setCellValue(contrib.getDonation(i).getAmount());
        }

        //format cells
        //Name cell
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
        final CellStyle leftBoldUnderline14Style = wb.createCellStyle();
        final Font boldUnderline14Font = wb.createFont();
        boldUnderline14Font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        boldUnderline14Font.setUnderline(Font.U_SINGLE);
        boldUnderline14Font.setFontHeightInPoints((short) 14);
        boldUnderline14Font.setFontName("Garamond");
        leftBoldUnderline14Style.setFont(boldUnderline14Font);
        leftBoldUnderline14Style.setAlignment(CellStyle.ALIGN_LEFT);
        cell[0][0].setCellStyle(leftBoldUnderline14Style);
        //field name cells
        final CellStyle rightBold10Style = wb.createCellStyle();
        final Font bold10Font = wb.createFont();
        bold10Font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        bold10Font.setFontHeightInPoints((short) 10);
        bold10Font.setFontName("Garamond");
        rightBold10Style.setFont(bold10Font);
        rightBold10Style.setAlignment(CellStyle.ALIGN_RIGHT);
        for (int i = 3; i < 22; i++) {
            cell[0][i].setCellStyle(rightBold10Style);
        }
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 0, 1));
        for (int i = 3; i < 11; i++) {
            cell[3][i].setCellStyle(rightBold10Style);
        }
        cell[4][0].setCellStyle(rightBold10Style);
        //field content cells
        final CellStyle left10Style = wb.createCellStyle();
        final Font garamond10Font = wb.createFont();
        garamond10Font.setFontHeightInPoints((short) 10);
        garamond10Font.setFontName("Garamond");
        left10Style.setFont(garamond10Font);
        left10Style.setAlignment(CellStyle.ALIGN_LEFT);
        for (int i = 3; i < 5; i++) {
            cell[1][i].setCellStyle(left10Style);
        }
        //phone number cells
        final CellStyle phoneStyle = wb.createCellStyle();
        phoneStyle.setFont(garamond10Font);
        phoneStyle.setAlignment(CellStyle.ALIGN_LEFT);
        final CreationHelper createHelper = wb.getCreationHelper();
        phoneStyle.setDataFormat(createHelper.createDataFormat().getFormat("[<=9999999]###-####;(###) ###-####"));
        for (int i = 5; i < 9; i++) {
            cell[1][i].setCellStyle(phoneStyle);
            sheet.addMergedRegion(new CellRangeAddress(i, i, 1, 2));

        }
        cell[1][9].setCellStyle(left10Style);
        //email through past contact
        for (int i = 10; i < 16; i++) {
            cell[1][i].setCellStyle(left10Style);
        }
        //notes
        CellStyle noteStyle = wb.createCellStyle();
        noteStyle.cloneStyleFrom(left10Style);
        noteStyle.setWrapText(true);
        cell[1][17].setCellStyle(noteStyle);
        //column E
        for (int i = 3; i < 11; i++) {
            cell[4][i].setCellStyle(left10Style);
        }
        //VanID Cell
        final CellStyle right10Style = wb.createCellStyle();
        right10Style.setFont(garamond10Font);
        right10Style.setAlignment(CellStyle.ALIGN_RIGHT);
        cell[5][0].setCellStyle(right10Style);
        //Notes cell
        sheet.addMergedRegion(new CellRangeAddress(17, 19, 1, 5));
        //contribution cells
        final CellStyle date10Style = wb.createCellStyle();
        date10Style.setFont(garamond10Font);
        date10Style.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
        date10Style.setBorderBottom(CellStyle.BORDER_THIN);
        date10Style.setBorderTop(CellStyle.BORDER_THIN);
        date10Style.setBorderLeft(CellStyle.BORDER_THIN);
        date10Style.setBorderRight(CellStyle.BORDER_THIN);
        final CellStyle contributionStyle = wb.createCellStyle();
        contributionStyle.cloneStyleFrom(left10Style);
        contributionStyle.setBorderBottom(CellStyle.BORDER_THIN);
        contributionStyle.setBorderTop(CellStyle.BORDER_THIN);
        contributionStyle.setBorderLeft(CellStyle.BORDER_THIN);
        contributionStyle.setBorderRight(CellStyle.BORDER_THIN);
        final CellStyle money10Style = wb.createCellStyle();
        money10Style.setFont(garamond10Font);
        money10Style.setDataFormat(createHelper.createDataFormat().getFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"));
        money10Style.setBorderBottom(CellStyle.BORDER_THIN);
        money10Style.setBorderTop(CellStyle.BORDER_THIN);
        money10Style.setBorderLeft(CellStyle.BORDER_THIN);
        money10Style.setBorderRight(CellStyle.BORDER_THIN);
        for (int i = 22; i < 22 + contrib.getDonationsLength(); i++) {
            cell[0][i].setCellStyle(date10Style);                
            cell[1][i].setCellStyle(contributionStyle);
            cell[2][i].setCellStyle(contributionStyle);
            cell[3][i].setCellStyle(contributionStyle);
            cell[4][i].setCellStyle(contributionStyle);
            sheet.addMergedRegion(new CellRangeAddress(i,i,1,4));
            cell[5][i].setCellStyle(money10Style);

            
            
        }
        //resize columns
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        try {
            FileOutputStream fileOut = new FileOutputStream(file);
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            return 1;
        } catch (IOException ex) {
            return 1;
        }

        return 0;
    }
}
