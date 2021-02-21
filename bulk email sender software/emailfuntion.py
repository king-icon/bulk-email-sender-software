import smtplib
def email_send_function(to_,subj_,msg_,from_,pass_):

    s=smtplib.SMTP("smtp.gmail.com",587) #create seesion
    s.starttls()  
    #transportlayer
    s.login(from_,pass_)

    msg = "subject: {}/n/n{}".format(subj_,msg_)
    s.sendmail(from_,to_,msg)
    x = s.ehlo()
    if x[0]==250:
        return "s"
    else:
        return "f"
    s.close()