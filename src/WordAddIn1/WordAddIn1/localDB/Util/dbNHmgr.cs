using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NHibernate;
using NHibernate.Cfg;


namespace OfficeAssist.localDB.Util
{
    public class dbNHmgr
    {

        private static readonly ISessionFactory sessionFactory;

        private static string HibernateHbmXmlFileName = "hibernate.cfg.xml";

        //private static ISession session

        static dbNHmgr()
        {
            String strCfgPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "hibernate.cfg.xml";

            Configuration cfg = new Configuration();

            Configuration cfgSub = cfg.Configure(strCfgPath);

            sessionFactory = cfgSub.BuildSessionFactory();

            // sessionFactory = new Configuration().Configure(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "hibernate.cfg.xml").BuildSessionFactory();

        }

        public static ISessionFactory getSessionFactory()
        {

            return sessionFactory;

        }

        public static ISession getSession()
        {

            return sessionFactory.OpenSession();

        }

        public static void closeSessionFactory()
        {
            return;
        }

    }
}
