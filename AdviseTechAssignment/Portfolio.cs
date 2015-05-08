using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdviseTechAssignment
{
    public class Portfolio
    {
        /// <summary>
        /// 
        /// </summary>
        private string name;
        public string Name { get { return name; } set { this.name = value; } }

        private string code;
        public string Code { get { return code; } set { this.code = value; } }

        private string market_value;
        public string MarketValue { get { return market_value; } set { this.market_value = value; } }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="n"></param>
        /// <param name="c"></param>
        /// <param name="mv"></param>
        public Portfolio(string n, string c, string mv)
        {
            this.Name = n;
            this.Code = c;
            this.MarketValue = mv;
        }

        /// <summary>
        /// shows instance variables of portoflio instance in a single string
        /// </summary>
        /// <returns>string of instance variables of a portfolio instance along with corresponding labels</returns>
        public override string ToString()
        {
            return "Name=" + this.Name + "\tCode=" + this.Code + "\tMarketValue=" + this.MarketValue;
        }
    }

    public class PortfolioCollection
    {
        /// <summary>
        /// instance variable is our List of Portfolio instances
        /// </summary>
        private List<Portfolio> portfolio_entities;
        public List<Portfolio> PortfolioEntities { get { return portfolio_entities; } set { this.portfolio_entities = value; } }

        /// <summary>
        /// constructor for PortfolioCollection instances
        /// </summary>
        public PortfolioCollection()
        {
            this.PortfolioEntities = new List<Portfolio>();
        }
    }
}
