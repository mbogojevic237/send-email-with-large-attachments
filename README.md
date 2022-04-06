# send-email-with-large-attachments

Trying to send emal using ms-graph API with attachments larger than 3MB will require usage of UploadSesion  (https://docs.microsoft.com/en-us/graph/outlook-large-attachments?tabs=http).

First of all, authorization is required (MSAL library or any other way).

> Note: ` Mail.ReadWrite permission` is minimum required.


## Step 1. Create new draft mail

```python
header = {
               'Authorization': 'Bearer ' + access_token,
               'Content-type': 'application/json',
               'Content-Length': '0'
}
```
In the body specifie mail parameters. See more https://docs.microsoft.com/en-us/graph/api/resources/message?view=graph-rest-1.0

```python
msg_new = {
    "subject":"TEST",
    "body":{
        "contentType":"HTML",
        "content":content
    },
    "toRecipients":[
        {
            "emailAddress":{
                "address":"wwe@xxx.com",
            }

        },
        {
            "emailAddress": {
                "address": "sdsds@xxx.com",
            }

        }
    ]
}
```
Create mail and get mail id for further usage.
```python
url = f'https://graph.microsoft.com/v1.0/users/yyyyyy@rrrrr.com/messages'
request_message = requests.post(url = url, headers=header, data = json.dumps(msg_new))
print(request_message.raise_for_status())

msg_id = request_message.json()['id']

```
## Step 2. Create UploadSession

Check file size.

``` python

f = os.stat('path to file\file')
size = f.st_size

```
Create UploadSession

``` python
header = {'Authorization': 'Bearer ' + access_token,  'Content-type': 'application/json'}
url = f'https://graph.microsoft.com/v1.0/users/yyyyyy@rrrrr.com/messages/{msg_id}/attachments/createUploadSession'

data ={
      "AttachmentItem": {
        "attachmentType": "file", 
        "name": filename,
        "size": size
      }
    }

req = requests.post(url=url, headers=header, data=json.dumps(data))
upload_url = req.json()['uploadUrl']
```
Set chunk size and determine how many chunks you will have based on file size.

```python
CHUNK_SIZE = 3500000
chunks = int(size / CHUNK_SIZE) + 1 if size % CHUNK_SIZE > 0 else 0
```
Start uploading the file.

```python
file = open(path\file, 'rb')
for i in range(chunks):
        chunk = file.read(CHUNK_SIZE)
        bytes_read = len(chunk)
        upload_range = f'bytes {start}-{start + bytes_read - 1}/{size}'

        attachment = requests.put(
            upload_url,
            headers={
                'Content-Type': 'application/json',
                'Content-Length': str(bytes_read),
                'Content-Range': upload_range
            },
            data=chunk
        )
        attachment.raise_for_status()
        start += bytes_read

        if chunk_num == chunks-1:

            t = attachment.headers
            location = t['Location'].split('/')
            
            for i in location:
                if 'users' in i:
                    user = i.replace('users', '').replace('(', '').replace(')', '')
                if 'messages' in i:
                    messages_id = i.replace('messages', '').replace('(', '').replace(')', '')
                if 'Attachments' in i:
                    attachmants_id = i.replace('Attachments', '').replace('(', '').replace(')', '')

```

## Step 3. Send mail

Now you have mail and attached file on that mail. Next step would be to send it.

``` python
url = f'https://graph.microsoft.com/v1.0/users/yyyyyy@rrrrr.com/messages/{msg_id}/send'
r = requests.post(url, headers=header)
print(r.raise_for_status())

```

## Addition

If the file size is less than 3MB, or you have to add additional small attachments on previous created mails:

``` python
url = f'https://graph.microsoft.com/v1.0/users/yyyyyy@rrrrr.com/messages/{msg_id}/attachments

data =  {

            "@odata.type":"#microsoft.graph.fileAttachment",
            "name":"test1.png",
            "isInline": True,
            "contentBytes": str(file.decode('utf8')), #super important to be base64
            'contentId' : 'image1'
          }

data2 =   {

    "@odata.type":"#microsoft.graph.fileAttachment",
    "name":"EMAIL List Update.xlsx",
    "isInline": False,
    "contentBytes": str(file.decode('utf8')),
    # 'contentId' : 'image1'
  }
  
requests.post(url=url, headers=header, data=json.dumps(data))
requests.post(url=url, headers=header, data=json.dumps(data2))

```

After this send mail like in Step 3.




















