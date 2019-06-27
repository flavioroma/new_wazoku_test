'''
We have requests from customers to send their admins a data export
of all user activity on the site.

This script queries the database for a list of all users and the content
they've interacted with on the site. It then constructs an xlsx file and then sends it
to an email address specified via a command line parameter
'''

import argparse
import csv
import os
from collections import Counter

from scripts.utils import Email

import django
django.setup()

from central import models
import itertools


os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'exercise.settings')

filename = 'users.csv'


def main(client, email):
    # Obtain all the entities we need to process
    site = models.Site.objects.get(domain=client)
    users = models.User.objects.filter(site=site)
    user_data = users.values_list('id', 'email')
    ideas = models.Idea.objects.filter(site=site)
    idea_creators_ids = ideas.values_list('creator_id', flat=True)
    challenges = models.Challenge.objects.filter(site=site)
    challenge_creators_ids = challenges.values_list('creator_id', flat=True)
    communities = models.Community.objects.filter(site=site)
    community_creators_ids = communities.values_list('creator_id', flat=True)
    conversations = models.Conversation.objects.filter(site=site)
    conversation_creators_ids = conversations.values_list('creator_id', flat=True)
    idea_comments = models.IdeaComment.objects.filter(site=site)
    idea_comment_creators_ids = idea_comments.values_list('creator_id', flat=True)
    idea_comments_liked_ids = []

    for comment in idea_comments:
        idea_comments_liked_ids.extend(
            comment.liked_by.all().values_list('id', flat=True)
        )

    idea_voters_ids = []

    for idea in ideas:
        idea_voters_ids.extend(
            idea.votes.all().values_list('creator_id', flat=True)
        )

    stage_comments_likers_ids = []

    idea_creators_counter = Counter(idea_creators_ids)
    challenge_creators_counter = Counter(challenge_creators_ids)
    idea_comment_counter = Counter(idea_comment_creators_ids)
    idea_comment_likes_counter = Counter(idea_comments_liked_ids)
    idea_votes_counter = Counter(idea_voters_ids)
    conversation_counter = Counter(conversation_creators_ids)
    community_counter = Counter(community_creators_ids)
    stage_comment_likes_counter = Counter(stage_comments_likers_ids)

    with open(filename, 'w') as csvfile:
        userwriter = csv.writer(csvfile)
        userwriter.writerow(['Active Users', 'Activities'])

        for id, user_email in user_data:
            activity = []

            if challenge_creators_counter[id]:
                action = 'The user created {} challenges.'.format(
                    challenge_creators_counter[id],
                )
                activity.append(action)

            if idea_creators_counter[id]:
                action = 'The user created {} ideas.'.format(
                    idea_creators_counter[id],
                )
                activity.append(action)

            if idea_comment_counter[id]:
                action = 'The user submitted {} idea comments.'.format(
                    idea_comment_counter[id],
                )
                activity.append(action)

            if idea_comment_likes_counter[id]:
                action = 'The user liked {} comments on ideas.'.format(
                    idea_comment_likes_counter[id],
                )
                activity.append(action)

            if idea_votes_counter[id]:
                action = 'The user voted {} ideas.'.format(
                    idea_votes_counter[id],
                )
                activity.append(action)

            if conversation_counter[id]:
                action = 'The user created {} conversations.'.format(
                    conversation_counter[id],
                )
                activity.append(action)

            if (community_counter[id]):
                action = 'The user created {} communities'.format(community_counter[id])
                activity.append(action)

            for community in communities:
                if community.managers.filter(id=id):
                    action = 'The user is a manager of the community "{}".'.format(
                        community.name,
                    )
                    activity.append(action)

                if community.members.filter(id=id):
                    action = 'The user is a member of the community "{}".'.format(
                        community.name,
                    )
                    activity.append(action)

            if stage_comment_likes_counter[id]:
                action = 'The user liked {} review comments.'.format(
                    stage_comment_likes_counter[id],
                )
                activity.append(action)

            if activity:
                userwriter.writerow([user_email] + activity)

    # Before sending, check we are only sending to a manager
    recipient = models.User.objects.get(email=email)
    if recipient.is_manager:
        email_sender = Email()
        email_sender.send_email(
            send_to=email,
            subject='User Activity Export',
            attachment=filename
        )

    # Clean up after ourselves
    os.remove(filename)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description='A script to generate and email an xlsx file containing user activity'
    )
    parser.add_argument(
        '-c', '--customer',
        help='the customer site',
        required=True,
    )
    parser.add_argument(
        '-e', '--email',
        help='the email address where to send the email with the output file',
        required=True,
    )

    return parser.parse_args()


if __name__ == "__main__":
    ns = parse_arguments()

    main(ns.customer, ns.email)
