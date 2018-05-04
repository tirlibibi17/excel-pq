Initially posted to /r/excel: [Querying the Reddit API with Power Query : excel](https://www.reddit.com/r/excel/comments/81y21n/querying_the_reddit_api_with_power_query/)

Here's a little project that queries the Reddit API recursively to fetch the \~1000 latest posts in /r/excel and filter them to show only the ones that are more than 2 days old and have been marked solved, but not by /u/Clippy_Office_Asst.

The goal of this is to be able to award a ClippyPoint for those posts, if [applicable](https://redd.it/7nexvy). The result looks like this: https://i.imgur.com/aNMKbUE.png.

**Some explanations**

The Reddit API pages the returned data and will only return a maximum of 25 items at a time, along with an id called "after" that you pass as an argument to your next call to get the next page. When after is returned null, there is no more data to return. 

The file built using 3 functions:

* UnixTime2Excel is just a utility function that converts a Unix time to Excel datetime format
* subreddit_new will fetch the 25 posts given the subreddit name and the after parameter
* SubredditNewRecursive is a recursive wrapper that builds the table of results recursively

And finally, the "points not awarded" query filters the results of SubRedditNewRecursive to keep only the posts whose link_flair_css_class is "solvedcase" and that are older than 50 hours. This is to comply with the "answered more than 2 days ago" rule, but it's approximate because:

* it compares local time (CET in my case) with UTC created time. Add or subtract hours based on your time zone.
* the correct answer may not have been provided in the hour that follows the post (although it often is)

Note that the functions can easily be adapted to call other Reddit API methods.