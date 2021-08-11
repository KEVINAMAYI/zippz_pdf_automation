import logging
import boto3
from botocore.exceptions import ClientError

BUCKET_NAME = "zippzpdfs"
s3 = boto3.resource('s3')


def generate_presigned_urls(uuid):
    try:
        logging.info("Start generating signed urls for uuid {}".format(uuid));
        s3 = boto3.client('s3')
        insert_file_url = s3.generate_presigned_url('get_object',
                                                    Params={'Bucket': BUCKET_NAME,
                                                            'Key': "{}/inserts.pdf".format(uuid)},
                                                    ExpiresIn=1600000)
        cards_file_url = s3.generate_presigned_url('get_object',
                                                    Params={'Bucket': BUCKET_NAME,
                                                            'Key': "{}/cards.pdf".format(uuid)},
                                                    ExpiresIn=1600000)
        logging.info("Finished generating signed urls for uuid {}".format(uuid));
        return insert_file_url, cards_file_url
    except ClientError as e:
        logging.error("Error generating presigned url {}".format(e));
        return None


def upload_file_to_aws(file_path,folder):
   file_name = file_path.split('/')[-1]
   logging.info("Uploading file with name {} started".format(file_path))
   try:
        s3.Bucket('zippzpdfs').upload_file(file_path, '{}/{}'.format(folder,file_name))

   except Exception as e:
       logging.error("Error uploading file with name {}".format(file_path))



def upload_files_to_aws(files_list,folder):
    for file_path in files_list:
        upload_file_to_aws(file_path,folder)
        logging.info("Uploading file with name {} finished".format(file_path))



#generate_presigned_url()