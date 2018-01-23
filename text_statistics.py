'''
Author : Ross Cole-Hunter (RVCHdotnet)
Description: Recursively searches folders looking for .docx files
    and analyses the text for unique word count, amongst other things.
    Results can then be visualised with the follow site:
        https://worditout.com/word-cloud/create

'''



import os
import docx2txt # Allows parsing of .docx word files to plain text
# not in the standard library, must be installed separate - see below
# pip install docx2txt
import re

blog_directory = '/Users/PEKOiSM/Documents/Euro Road Trip Blog/'
#TODO - replace above with sys.argv[1]
blog_post_list = [] # list of all the .docx filenames found
blog_post_length = [] # list of all the posts word counts
blog_words = {} # dict of all the unique words and occurence
min_char = 4 # minimum character length of words considered - i.e. ignore below
#TODO - replace above with sys.argv[2]
max_unique = 1000 # output the top max_unique words
#TODO - also take as sys.argv[3]

# Iterate through source directory searching for all .docx word documents
# Return list of all docx file names (including full path)
for root, dirnames, filenames in os.walk(blog_directory):
    for filename in filenames:
        if filename.endswith('docx'):
            blog_post_list.append(os.path.join(root, filename))

# Analyse each of the blog posts
# Return a dict with words/counts of most popular words used
for filename in blog_post_list:
    word_count = 0  # temporary count of words in post, saves re-iterating
    blog_post = docx2txt.process(filename)
    for word in blog_post.lower().split():
        word_count += 1
        # remove all punctiation from word
        word = re.sub(r'[^\w\s]', '', word)
        # ignore the words less than min_char letters long
        if len(word) > min_char:
            # add new words into dict, increment value for existing keys
            if word not in blog_words:
                blog_words[word] = 1
            else:
                blog_words[word] += 1
    # print 'The post {0} has {1} words.'.format(filename, word_count)
    #TODO - print the above if flag passed as argument
    blog_post_length.append(word_count)

# create a new list, sorted by frequency of occurence, limited to top 150
blog_words_sorted = sorted(blog_words.items(), key=lambda x: x[1], reverse=True)[:max_unique]

print 'The {0} most commonly used words greater than '\
        '{1} characters are:'.format(max_unique, min_char)
print '{:<4} {:<12} {:>8}'.format('Index', 'Word', 'Frequency')
i = 1
for k,v in blog_words_sorted:
        print '{:<4} {:<12} {:>8}'.format(i, k, v)
        i += 1

# TODO - print the below if flag requests it
# Count and Return the number of blog posts
# print 'There are %s blog posts in %s' %(len(blog_post_list), blog_directory)
print 'There are {0} blog posts in {1}'.format(len(blog_post_list), blog_directory)

# Count the word count of all posts
# Return the Average, the Max and the Min
print 'There are {0} unique words used.'.format(len(blog_words))
print 'There are {0} total words used.'.format(sum(blog_post_length))
print 'The average post has {0} words.'.format(sum(blog_post_length)/len(blog_post_length))
print 'The maximum post uses {0} words.'.format(max(blog_post_length))
print 'The minimum post uses {0} words.'.format(min(blog_post_length))
