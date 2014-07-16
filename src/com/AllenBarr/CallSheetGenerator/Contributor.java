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

import java.math.BigInteger;
import java.util.ArrayList;

/**
 *
 * @author captainbowtie
 */
public class Contributor {
    private Integer vanID=0;
    private String name="";
    private String salutation="";
    private String sex="";
    private String party="";
    private Double phone = 0.0;
    private Double homePhone=0.0;
    private Double cellPhone=0.0;
    private Double workPhone=0.0;
    private String email="";
    private Double workExtension=0.0;
    private String employer="";
    private String occupation="";
    private Integer age=0;
    private String spouse="";
    private String streetAddress="";
    private String city="";
    private String state="";
    private Integer zip=0;
    private String notes="";
    private ArrayList<Donation> donations = new ArrayList<Donation>();
    public int getDonationsLength(){
        return donations.size();
    }

    public Integer getVANID() {
        return vanID;
    }

    public String getName() {
        return name;
    }

    public String getSalutation() {
        return salutation;
    }

    public String getSex() {
        return sex;
    }

    public String getParty() {
        return party;
    }

    public Double getPhone() {
        return phone;
    }

    public Double getHomePhone() {
        return homePhone;
    }

    public Double getCellPhone() {
        return cellPhone;
    }

    public Double getWorkPhone() {
        return workPhone;
    }

    public Double getWorkExtension() {
        return workExtension;
    }

    public String getEmail() {
        return email;
    }
    
    public String getEmployer() {
        return employer;
    }

    public String getOccupation() {
        return occupation;
    }

    public Integer getAge() {
        return age;
    }

    public String getSpouse() {
        return spouse;
    }

    public String getStreetAddress() {
        return streetAddress;
    }

    public String getCity() {
        return city;
    }

    public String getState() {
        return state;
    }

    public Integer getZip() {
        return zip;
    }

    public String getNotes() {
        return notes;
    }

    public ArrayList<Donation> getDonations() {
        return donations;
    }
    
    public Donation getDonation(Integer i){
        return donations.get(i);
    }

    public void setVANID(Integer vanID) {
        this.vanID = vanID;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setSalutation(String salutation) {
        this.salutation = salutation;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public void setParty(String party) {
        this.party = party;
    }

    public void setPhone(Double phone) {
        this.phone = phone;
    }

    public void setHomePhone(Double homePhone) {
        this.homePhone = homePhone;
    }

    public void setCellPhone(Double cellPhone) {
        this.cellPhone = cellPhone;
    }

    public void setWorkPhone(Double workPhone) {
        this.workPhone = workPhone;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public void setWorkExtension(Double workExtension) {
        this.workExtension = workExtension;
    }

    public void setEmployer(String employer) {
        this.employer = employer;
    }

    public void setOccupation(String occupation) {
        this.occupation = occupation;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public void setSpouse(String spouse) {
        this.spouse = spouse;
    }

    public void setStreetAddress(String streetAddress) {
        this.streetAddress = streetAddress;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public void setState(String state) {
        this.state = state;
    }

    public void setZip(Integer zip) {
        this.zip = zip;
    }

    public void setNotes(String notes) {
        this.notes = notes;
    }

    public void addDonation(Donation d){
        this.donations.add(d);
    }

    
}
