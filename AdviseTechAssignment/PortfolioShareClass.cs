using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdviseTechAssignment
{
    public class PortfolioShareClass
    {
        /// <summary>
        /// instance variables
        /// </summary>
        private string parent_portfolio;
        public string ParentPortfolio { get { return parent_portfolio; } set { this.parent_portfolio = value; } }

        private string portfolio_share_class_name;
        public string PortfolioShareClassName { get { return portfolio_share_class_name; } set { this.portfolio_share_class_name = value; } }

        private string portfolio_share_class_code;
        public string PortfolioShareClassCode { get { return portfolio_share_class_code; } set { this.portfolio_share_class_code = value; } }

        private string portfolio_share_class_base_fee;
        public string PortfolioShareClassBaseFee { get { return portfolio_share_class_base_fee; } set { this.portfolio_share_class_base_fee = value; } }

        /// <summary>
        /// constructor for the Portfolio Share Class
        /// </summary>
        /// <param name="pp">Parent Portfolio</param>
        /// <param name="pscn">Portfolio Share Class Name</param>
        /// <param name="pscc">Portfolio Share Class Code</param>
        /// <param name="pscbf">Portfolio Share Class Base Fee</param>
        public PortfolioShareClass(string pp, string pscn, string pscc, string pscbf)
        {
            this.ParentPortfolio = pp;
            this.PortfolioShareClassName = pscn;
            this.PortfolioShareClassCode = pscc;
            this.PortfolioShareClassBaseFee = pscbf;
        }

        /// <summary>
        /// shows instance variables of portoflio share class instance in a single string
        /// </summary>
        /// <returns>string of instance variables of a portfolio instance along with corresponding labels</returns>
        public override string ToString()
        {
            return "ParentPortfolio=" + this.ParentPortfolio + "\tShareClassName=" + this.PortfolioShareClassName +
                "\tShareClassCode=" + this.PortfolioShareClassCode + "\tShareClassBaseFee=" + this.PortfolioShareClassBaseFee;
        }
    }

    public class PortfolioShareClassCollection
    {
        /// <summary>
        /// instance variable is our List of PortfolioShareClass instances. 
        /// </summary>
        private List<PortfolioShareClass> portfolio_share_class_entities;
        public List<PortfolioShareClass> PortfolioShareClassEntities { get { return portfolio_share_class_entities; } set { this.portfolio_share_class_entities = value; } }

        /// <summary>
        /// constructor for a PortfolioShareClassCollection
        /// </summary>
        public PortfolioShareClassCollection()
        {
            this.PortfolioShareClassEntities = new List<PortfolioShareClass>();
        }
    }
}
