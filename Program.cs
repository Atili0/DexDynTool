using System.Collections;

namespace DeXrm.Console
{
    using System;
    using System.Linq;
    using System.Security;
    using Microsoft.Xrm.Tooling.Connector;
    using System.Collections.Generic;
    using BizArk.Core.CmdLine;


    class Program
    {
        [CmdLineDefaultArg("Help")]
        internal class DeXrmCmdLineObject
       : CmdLineObject
        {
            [CmdLineArg(Alias = "o")]
            [System.ComponentModel.Description("Nombre de la organizacion")]
            public string Org { get; set; }

            [CmdLineArg(Alias = "s")]
            [System.ComponentModel.Description("Nombre o ip del servidor")]
            public string Server { get; set; }

            [CmdLineArg(Alias = "p")]
            [System.ComponentModel.Description("Puerto del Servidor")]
            public string Port { get; set; }

            [CmdLineArg(Alias = "u")]
            [System.ComponentModel.Description("Usuario del CRM")]
            public string User { get; set; }
        }

        static void Main(string[] args)
        {
            ChangeConsoleColor();

            if (args.Length > 0 && args[0].Equals(@"\?"))
            {
                DeXrmCmdLineObject target;
                target = new DeXrmCmdLineObject();
                target.InitializeEmpty();
                Console.WriteLine(target.GetHelpText(50));
                Console.Read();
            }
            else
            {
                Connection.CrmSvc = new Program().GetConnection();
                Util.ColoredConsoleWrite(ConsoleColor.DarkGreen, "Que acción desea realizar");
                Console.WriteLine("1- Crear una entidad");
                Console.WriteLine("2- Crear campos de una entidad");
                var selec = Console.ReadLine();
                switch (selec)
                {
                    case "1":
                        new Program().GetInformationEntity(Connection.CrmSvc);
                        break;
                    case "2":
                        Util.ColoredConsoleWrite(ConsoleColor.DarkGreen, "Schema de la entidad :");
                        var entity = System.Console.ReadLine();
                        new Program().GetEntityAttribute(Connection.CrmSvc, entity);
                        break;
                }

            }
        }

        private static void ChangeConsoleColor()
        {
            Console.BackgroundColor = ConsoleColor.Blue;
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.White;
        }

        public void GetInformationEntity(CrmServiceClient service)
        {
            var pathfile = String.Format(@"{0}\{1}{2}", 
                AppDomain.CurrentDomain.BaseDirectory,
                @"Configuration\",
                System.Configuration.ConfigurationManager.AppSettings["File"]);

            List<DxEntity> lEntities = new Util().GetListEntity(Util.ReadExcel(pathfile));

            foreach (var entity in lEntities  )
            {
                new DynamicsCRM().CreateEntity(service, entity);
            }
        }

        public void GetEntityAttribute(CrmServiceClient service, string entity)
        {
            var pathfile = String.Format(@"{0}\{1}{2}",
               AppDomain.CurrentDomain.BaseDirectory,
               @"Configuration\",
               System.Configuration.ConfigurationManager.AppSettings["FileAttribute"]);

            List<DxEntityAttribute> lEntities = new Util().GetListAttribute(Util.ReadExcel(pathfile));

            new DynamicsCRM().CreateAttribute(service, entity, lEntities);
        }

        public CrmServiceClient GetConnection()
        {
            System.Console.WriteLine("Escriba el nombre del servidor :");
            var server = System.Console.ReadLine();
            System.Console.WriteLine("Escriba el puerto :");
            var port = System.Console.ReadLine();
            System.Console.WriteLine("Escriba el nombre de la organización");
            var org = System.Console.ReadLine();
            System.Console.WriteLine("Desea conectar al CRM con el usuario actual? Y/N");
            var readLine = System.Console.ReadLine();
            if (readLine != null && readLine.ToUpper().Equals("Y"))
            {
                Connection.CrmSvc = new CrmServiceClient(System.Net.CredentialCache.DefaultNetworkCredentials,
                    AuthenticationType.AD, server, port, org);
            }
            else
            {
                System.Console.WriteLine("Escriba el dominio del usuario :");
                var dominio = System.Console.ReadLine();
                System.Console.WriteLine("Escriba el usuario a conectar :");
                var usuario = System.Console.ReadLine();
                System.Console.WriteLine("Escriba el password del usuario :");
                var password = Util.ReadPassword();

                Connection.CrmSvc =
                    new CrmServiceClient(new System.Net.NetworkCredential(usuario, ToSecureString(password), dominio),
                        AuthenticationType.AD, server, port, org);
            }

            return Connection.CrmSvc;
        }
        private SecureString ToSecureString(string phrase)
        {
            SecureString secureString = new SecureString();
            phrase.ToCharArray().ToList().ForEach(p => secureString.AppendChar(p));

            return secureString;
        }


    }
}
