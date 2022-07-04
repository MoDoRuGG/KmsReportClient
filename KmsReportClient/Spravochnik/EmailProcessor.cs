using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;

namespace KmsReportClient.Spravochnik
{
    public class EmailProcessor
    {
        private readonly EndpointSoap _client;
        public EmailProcessor(EndpointSoap client)
        {
            _client = client;
        }

        public void AddEmail(string email, string description)
        {
            _client.AddEmail(new AddEmailRequest
            {
                Body = new AddEmailRequestBody
                {
                    description = description,
                    email = email
                }
            });


        }


        public void EditEmail(int emailId, string email, string description)
        {
            _client.EditEmail(new EditEmailRequest
            {
                Body = new EditEmailRequestBody
                {
                    emailId = emailId,
                    description = description,
                    email = email
                }
            });
        }



        public void EditEmail(int emailId)
        {
            _client.DeleteEmail(emailId);
        }

    }
}
