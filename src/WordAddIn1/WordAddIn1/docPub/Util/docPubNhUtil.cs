using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NHibernate;
using NHibernate.Cfg;


namespace OfficeAssist.NH.Util
{
    public class docPubNhUtil
    {
        private static readonly ISessionFactory sessionFactory;
        // private static string HibernateHbmXmlFileName = "hibernate.cfg.xml";

        static docPubNhUtil()
        {
            sessionFactory = new Configuration().Configure().BuildSessionFactory();
            return;
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

    }// class

}// namespace
