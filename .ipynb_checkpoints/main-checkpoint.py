from pyChatGPT import ChatGPT
from mailmerge import MailMerge
import os
from datetime import datetime

COVER_LETTER='data/cover_letter_template.docx'
SESSION_TOKEN='eyJhbGciOiJkaXIiLCJlbmMiOiJBMjU2R0NNIn0..yY78KtFUuN2LdkqA.HLcUN499vB0-Ogz0mSDeNvuIVGSsTO4yG4HeBvYLJlRyC8yGYwt8LA-y9Bl-YpvZ-IP3VGjKGrTKYN4iLRKRl7DXDfvNhrbT0WDdy0FyIhao6DLaexmCznmNehOz6Y9d6KnLRI4LaMp_p8gnjVmuvI4LD20bSsD6z4KtkZJLdfQn3j58SM0rlWeQzNg1mw2sI_T6KQh12t-BylZYt5BuBlTgFaf_3Qloqb-Elh38C77ltBy-xkIRuGCrBWiiLdOjl7-NoH1fUOsiZk9_OfthYhhQ_k9g4RxJY2zairqz2lm4kIfs7gjKWBVHSxCtP0YOgbIr7hiBWPfJL-_Rx-oDrgL1crYN5402pvVP3VCkFzUzZ3vmNhENm5X8F0MfEiz9xPAdy6hYvWlpRnYV-S7bjTPKGIrCYqIFZlM3W8M_EsarchLzCN7TMK7hNIFhw5zlmU5K2DAxNGU3vLygdGieIyILTuL5M0sb4NKQm2PZxdNQqClGFVSrLpaAKIzaz_T-n4mOwaWCx3QYNIo8CAcN2vXgVL4PHOEUog9z-VxkupjNSHaJHqSPPsbxBH3eO0Wn4_L657Gu1I5QMPxOW7R7gtHKgP--4nMb_E8gf9VVytu7CTo_9fSkeqHmDtaWEXGUmhZakyR0kUp2sulvdLmRMWtlfeCoLiEVmG3lasz70KkfvTQ392wze4TmcqxA-A3nIH2JzeMGd9QZeYycqgf0ZggzbFK5vp7kECXgdBTOqTWZSvE1uSDGkhfZQr0jzO-eX5WpFS4lPP0UcAFyl89EV3zWJ2K-SN8anSv_zgsBMN-8TYEy7q7yrPsL7d2bgN-Hlj4eX5u5cK9Ol4fEQt5RbgP7FQzUzovYk-QIG0nyCKtrxppMgUGDwK5bOjJZbWRrk6mlCSTomyDQposRFCAe-GP3N1hbAD9XsnRgHne4Sb6ryjywGFjBel3bKGZ-3sCm3qMsDIb8rkMu_CBXflfRoUNHuhjHE2vnjvtGOdiJrEBmlotr3Ox130GVR7J6dhxHqKNfdetMwEYvS4Oas3qO4whnjQ5plAam6S3Spt186u6sWH0JhlZTRSySFNncr66oJk-1AqfJW4UGPSfL8teoslVUmUhjwn95e-MBRYMtkSgSjLPv7yMhqbtBNG_r0Q9Zv6KhHF5B0XzTa0nLd2q2xuoQ8meN1ddGf3R_pS7Rzu4btCR6PVZY2PTo-C4wRtPpZKr3Dgr1symF7JTauVe3kOxDfyjQ4Iq00aE1-E1Am9VbuG2y_yvAcXLt5xScFJAvBudLeA9DZ-JLOXKk6RlYLV6_OgwRUOJLJAEelxKV55Jua49wAZ88M4ytTDfR9n3lF5iSa63sRnnB-ScL5DT7GuLg1UaeozQCuYf3z4ZwlV2jVfkUVkNkm2awqhr7jpFB4iIVXPoLu3U8646-JoGzPGgXEsvWz3PAqrflowVdhjasQuHXKJowd0LfHk3LoZ1YNN3jC0NF3AfyNPuCLLXvtYYBbFMLavNmWmS94QEmYUU6H-2FNjo3T_km9uurHqtF9evfZm1ZaHPHDviyhg6C0SqEJ5qjqM3V_Uvm4Yl12us_FqBRDUZbUj3KAwlQBNnHWj9VT92_Z8knLdWOHzlu_RaoBg9GauNedvwgg7HUUgiCy3kQvA74GgHf6PwkmmF2RmrXxEJZ5SLdvga2KBNHJ7dZwHhCQPWTvDz8UJqP9mWQ8MSqsT_5x1PCsVEZVGc9h8zyd8cxyIKNu4M3EVqg7NFUS7PK6HkmZVJcDo4Jlh_hbzJdkrQs4l_iWliK6arVaUnGIA1B-s4iS0GYAIfr0ednxUPflozeYr8I47qGsY8V02Gy7KeI4WFP82sfIIl-MjIqDsArBIl4L8BDCWSZkibV93cOYbUQrBwOs_u_HBY1FnQsj2RvYMHeabU8F3qoi4vWTGPASiiVZDb3go2io5eR6apvjKDgmuJGZbG_DKuOIzb34UO_Bwv0J1bsdCa6f4IEoh-l5PVDuVhsxs1nPyeJuVxOdPNmmtTXC0BYSqx1QE9RxmyBTAsta_bypxT4VV6ijsZ8oLXqcCuApb27ZQOTapwSyb6qRTa8pcCxW2SO6CBXJa-z253--nXgeVIyp6EooVNYz3rnDixDCRT3gSbqk4ien9TMR7BlbiFjcrx7YYZ4Bgd4aGHNjPMbUUEBaYomFVnnBm22pLDb_eXDqVPgL52N0Td6bgSzAaq0ZkUzaxzv0F7t3ONMDfbZOuTDWQ6k9UWjs4TN73c23HaS-1HYSTGh1EMvKDOyR_de4nq05u1NexVokrcI12IoG8Cle8OceUyhzHlworuPg1fi8kfS6gXncQJBlSi6Nnbuw9cQdYzQRiRcRBRaEkwvaoMzz0JaWhAPs0PWEr1k5-5_ldIO-ovvJqyKimO1NEi0WFMh0ShEXtvAlR-1eCYFOqL2b-tQKlmGWyZpDfn5O99rB1NOaAVke4IuZfSxbrbvizbU3VJChVJewujRIqPukP_P48woizk6qFRON42QJWhIEfWiExBmOVDyV5uFwP_tRMKaZPxSFjdfaEBaHw.LumSyO3L7UWxSYe-L22prw'


def mail_merge(template):
    document=MailMerge(template)
    return document

def auth_chat_gpt():
    api=ChatGPT(SESSION_TOKEN,moderation=False,auth_type='openai', captcha_solver='2captcha', solver_apikey='abc')
    return api

def gpt_cover_letter(message):
    PROMT_COVER_LETTER=\
        """
            I want you to act as a recruiter. 
            Use this text and create a short sentence about why you like the company. 
            Make the sentence start from I become an admirer.
            "{company_description}"
        """.format(company_description=message)
    api=auth_chat_gpt()
    response=api.send_message(PROMT_COVER_LETTER)['message']
    return str(response).rstrip().lstrip().replace('"','')


def job_description():
    NAME = input("Company Name: ")
    LOCATION = input("Company Location: ")
    POSITION = input("Position: ")
    COMPANY_DESCRIPTION = input("Compnay Description: ")
    JOB_DESCRIPTION = input("Job Description: ")
    QUALIFICATION = input("Skills/Qualification: ")
    DATE=datetime.today().strftime('%B %d, %Y')
    
    container={
        'company':[
            {
            'name':NAME,
            'location':LOCATION,
            'position':POSITION,
            'background':COMPANY_DESCRIPTION
            }
        ],
        'description':JOB_DESCRIPTION,
        'skills':QUALIFICATION,
        'date':DATE
    }
    return container

def generate_cover_letter(container,template):
    document=mail_merge(template)
    message=container.get('company')[0]['background']
    document.merge(DATE=container.get('date'),
                   COMPANY_NAME=container.get('company')[0]['name'],
                   COMPANY_LOCATION=container.get('company')[0]['location'],
                   POSITION=container.get('company')[0]['position'],
                   MESSAGE=gpt_cover_letter(message))

    document.write('result.docx')

def generate_resume(template):
    pass


response=job_description()
print(response)
generate_cover_letter(response,COVER_LETTER)



# TEMPLATE='data/template.docx'
# document=MailMerge(TEMPLATE)
# field_name=(document.get_merge_fields())
# date=datetime.today().strftime('%B %d, %Y')
# name='Rajiv'

# print(datetime.today().strftime('%B %d, %Y'))

##session_token='eyJhbGciOiJkaXIiLCJlbmMiOiJBMjU2R0NNIn0..yY78KtFUuN2LdkqA.HLcUN499vB0-Ogz0mSDeNvuIVGSsTO4yG4HeBvYLJlRyC8yGYwt8LA-y9Bl-YpvZ-IP3VGjKGrTKYN4iLRKRl7DXDfvNhrbT0WDdy0FyIhao6DLaexmCznmNehOz6Y9d6KnLRI4LaMp_p8gnjVmuvI4LD20bSsD6z4KtkZJLdfQn3j58SM0rlWeQzNg1mw2sI_T6KQh12t-BylZYt5BuBlTgFaf_3Qloqb-Elh38C77ltBy-xkIRuGCrBWiiLdOjl7-NoH1fUOsiZk9_OfthYhhQ_k9g4RxJY2zairqz2lm4kIfs7gjKWBVHSxCtP0YOgbIr7hiBWPfJL-_Rx-oDrgL1crYN5402pvVP3VCkFzUzZ3vmNhENm5X8F0MfEiz9xPAdy6hYvWlpRnYV-S7bjTPKGIrCYqIFZlM3W8M_EsarchLzCN7TMK7hNIFhw5zlmU5K2DAxNGU3vLygdGieIyILTuL5M0sb4NKQm2PZxdNQqClGFVSrLpaAKIzaz_T-n4mOwaWCx3QYNIo8CAcN2vXgVL4PHOEUog9z-VxkupjNSHaJHqSPPsbxBH3eO0Wn4_L657Gu1I5QMPxOW7R7gtHKgP--4nMb_E8gf9VVytu7CTo_9fSkeqHmDtaWEXGUmhZakyR0kUp2sulvdLmRMWtlfeCoLiEVmG3lasz70KkfvTQ392wze4TmcqxA-A3nIH2JzeMGd9QZeYycqgf0ZggzbFK5vp7kECXgdBTOqTWZSvE1uSDGkhfZQr0jzO-eX5WpFS4lPP0UcAFyl89EV3zWJ2K-SN8anSv_zgsBMN-8TYEy7q7yrPsL7d2bgN-Hlj4eX5u5cK9Ol4fEQt5RbgP7FQzUzovYk-QIG0nyCKtrxppMgUGDwK5bOjJZbWRrk6mlCSTomyDQposRFCAe-GP3N1hbAD9XsnRgHne4Sb6ryjywGFjBel3bKGZ-3sCm3qMsDIb8rkMu_CBXflfRoUNHuhjHE2vnjvtGOdiJrEBmlotr3Ox130GVR7J6dhxHqKNfdetMwEYvS4Oas3qO4whnjQ5plAam6S3Spt186u6sWH0JhlZTRSySFNncr66oJk-1AqfJW4UGPSfL8teoslVUmUhjwn95e-MBRYMtkSgSjLPv7yMhqbtBNG_r0Q9Zv6KhHF5B0XzTa0nLd2q2xuoQ8meN1ddGf3R_pS7Rzu4btCR6PVZY2PTo-C4wRtPpZKr3Dgr1symF7JTauVe3kOxDfyjQ4Iq00aE1-E1Am9VbuG2y_yvAcXLt5xScFJAvBudLeA9DZ-JLOXKk6RlYLV6_OgwRUOJLJAEelxKV55Jua49wAZ88M4ytTDfR9n3lF5iSa63sRnnB-ScL5DT7GuLg1UaeozQCuYf3z4ZwlV2jVfkUVkNkm2awqhr7jpFB4iIVXPoLu3U8646-JoGzPGgXEsvWz3PAqrflowVdhjasQuHXKJowd0LfHk3LoZ1YNN3jC0NF3AfyNPuCLLXvtYYBbFMLavNmWmS94QEmYUU6H-2FNjo3T_km9uurHqtF9evfZm1ZaHPHDviyhg6C0SqEJ5qjqM3V_Uvm4Yl12us_FqBRDUZbUj3KAwlQBNnHWj9VT92_Z8knLdWOHzlu_RaoBg9GauNedvwgg7HUUgiCy3kQvA74GgHf6PwkmmF2RmrXxEJZ5SLdvga2KBNHJ7dZwHhCQPWTvDz8UJqP9mWQ8MSqsT_5x1PCsVEZVGc9h8zyd8cxyIKNu4M3EVqg7NFUS7PK6HkmZVJcDo4Jlh_hbzJdkrQs4l_iWliK6arVaUnGIA1B-s4iS0GYAIfr0ednxUPflozeYr8I47qGsY8V02Gy7KeI4WFP82sfIIl-MjIqDsArBIl4L8BDCWSZkibV93cOYbUQrBwOs_u_HBY1FnQsj2RvYMHeabU8F3qoi4vWTGPASiiVZDb3go2io5eR6apvjKDgmuJGZbG_DKuOIzb34UO_Bwv0J1bsdCa6f4IEoh-l5PVDuVhsxs1nPyeJuVxOdPNmmtTXC0BYSqx1QE9RxmyBTAsta_bypxT4VV6ijsZ8oLXqcCuApb27ZQOTapwSyb6qRTa8pcCxW2SO6CBXJa-z253--nXgeVIyp6EooVNYz3rnDixDCRT3gSbqk4ien9TMR7BlbiFjcrx7YYZ4Bgd4aGHNjPMbUUEBaYomFVnnBm22pLDb_eXDqVPgL52N0Td6bgSzAaq0ZkUzaxzv0F7t3ONMDfbZOuTDWQ6k9UWjs4TN73c23HaS-1HYSTGh1EMvKDOyR_de4nq05u1NexVokrcI12IoG8Cle8OceUyhzHlworuPg1fi8kfS6gXncQJBlSi6Nnbuw9cQdYzQRiRcRBRaEkwvaoMzz0JaWhAPs0PWEr1k5-5_ldIO-ovvJqyKimO1NEi0WFMh0ShEXtvAlR-1eCYFOqL2b-tQKlmGWyZpDfn5O99rB1NOaAVke4IuZfSxbrbvizbU3VJChVJewujRIqPukP_P48woizk6qFRON42QJWhIEfWiExBmOVDyV5uFwP_tRMKaZPxSFjdfaEBaHw.LumSyO3L7UWxSYe-L22prw'
##api=ChatGPT(session_token,moderation=False,auth_type='openai', captcha_solver='2captcha', solver_apikey='abc')


# PROMPT="""
#         I want you to act as a recruiter. 
#         Use this text and create strong bullet points from it. 
#         Give me top 3 responses with in-depth technical details.
#         "
#             {job}
#         "


#         """.format(name=name)
# print(PROMPT)
#response=api.send_message('What you think about blue origin')
# ##print(response['message'])
# # document.merge(date=DATE)
# # document.write('result.docx')


