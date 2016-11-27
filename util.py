# !/usr/bin/env python
# encoding:UTF-8
import urllib2
def request_url(url,repeat=3):
	ret = -1#失败
	content = ''
	for cnt in xrange(repeat):
		try:
			req = urllib2.Request(url);
			response = urllib2.urlopen(req)
			content = response.read()
			response.close()
			ret = 0#成功
			break;
		except:
			continue
	result = (ret,content)
	return result