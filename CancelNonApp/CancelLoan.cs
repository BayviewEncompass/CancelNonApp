using EllieMae.Encompass.Automation;
using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.BusinessObjects.Users;
using EllieMae.Encompass.Collections;
using EllieMae.Encompass.ComponentModel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CancelNonApp
{
    [Plugin]
    public class CancelLoan
    {

        public CancelLoan()
        {
            EncompassApplication.LoanOpened += EncompassApplication_LoanOpened;
            

            
        }

        private void EncompassApplication_LoanOpened(object sender, EventArgs e)
        {
          EncompassApplication.LoanClosing += EncompassApplication_LoanClosing;
        }

        private void EncompassApplication_LoanClosing(object sender, EventArgs e)
        {
            User user = EncompassApplication.CurrentUser;
            if (user.ID == "christophercl")
            {
                Loan loan = EncompassApplication.CurrentLoan;
                if (loan.Fields["CX.REQUEST.CANCEL.NOAPP"].Value.ToString() == "X" & loan.Fields["3142"].Value == null)
                {
                    loan.Session.Loans.Folders["Cancelled No Application"].IsTrash.Equals(true);
                    loan.MoveToFolder(loan.Session.Loans.Folders["Cancelled No Application"]);
                }
            }
        }
    }
}
