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

import java.util.Date;



/**
 *
 * @author captainbowtie
 */
public class Donation {
    private final Date donationDate;
    private final String recipient;
    private final Double amount;
    public Donation(final Date date, final String name, final Double amt){
        donationDate=date;
        recipient=name;
        amount=amt;
    }

    public Date getDonationDate() {
        return donationDate;
    }

    public String getRecipient() {
        return recipient;
    }

    public Double getAmount() {
        return amount;
    }
    
}
