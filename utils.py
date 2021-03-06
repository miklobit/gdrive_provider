import unicodedata
import re


def slugify (s):
    if type(s) is unicode:
        slug = unicodedata.normalize('NFKD', s)
    elif type(s) is str:
        slug = s
    else:
        raise AttributeError("Can't slugify string")
    slug = slug.encode('ascii', 'ignore').lower()
    slug = slug.replace(' ', '_')
    #slug = re.sub(r'[^a-z0-9]+', '-', slug).strip('-')
    #slug = re.sub(r'--+',r'-',slug)
    return slug