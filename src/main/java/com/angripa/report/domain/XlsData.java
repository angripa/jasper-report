package com.angripa.report.domain;

public class XlsData {
   private String date;
   private String detailTrx;
   private String teller;
   private Double debit;
   private Double credit;
   private Double balance;

   public String getDate() {
      return date;
   }

   public void setDate(String date) {
      this.date = date;
   }

   public String getDetailTrx() {
      return detailTrx;
   }

   public void setDetailTrx(String detailTrx) {
      this.detailTrx = detailTrx;
   }

   public String getTeller() {
      return teller;
   }

   public void setTeller(String teller) {
      this.teller = teller;
   }

   public Double getDebit() {
      return debit;
   }

   public void setDebit(Double debit) {
      this.debit = debit;
   }

   public Double getCredit() {
      return credit;
   }

   public void setCredit(Double credit) {
      this.credit = credit;
   }

   public Double getBalance() {
      return balance;
   }

   public void setBalance(Double balance) {
      this.balance = balance;
   }
}
