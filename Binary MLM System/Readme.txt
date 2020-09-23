DESCRIPTION

This is a genealogy report generation for a binary structured multi-level marketing system.

The report is in listview, but stored in an array. The program is a little slow but there are
other ways to process the report faster. The program took 1 min. to process more than 600 members
(at 31 levels). It will work a lot slower if the all the nodes in the binary tree is full
(have 2 downlines each). You can use the binary search in searching the array to make it
much faster because i used only sequential search.

You have to figure out how to make a printable version of the report. ;-)

GENEALOGY REPORT

See Genea Illustrations folder.

BINARY MLM SYSTEM

A binary mlm system has a binary tree structure where a node (or member) has only 1 or 2
(Left and/or Right) leg or downline nodes. A node is connected to an upline node, the topmost
node or is the only node that is not connected to an upline node. The main idea is that a
node or a member has one or two members under him and his one or two members has also one
or two downlines each, and so on. The binary structure grows downward as more nodes are added
to it.

A member has two wings; the left and right wing. A member's left wing is composed of all nodes
that are connected to his left node regardless of the nodes' binary position underneath his
left node. A member's right wing is composed of all nodes that are connected to his right node
regardless of the nodes' binary position underneath his right node.

Example:

(M) is the topmost node in the binary tree. (M) has two downlines: (N) and (O). (N) is
connected to the left side of (M) and (O) is connected to the right side of (M). N has two
downlines: (P) and (Q). (P) is connected to the left side of (N) and (Q) is connected to the
right side of (N). (O) has only one downline: (R). (R) is connected to the left side of (O).

          Sample Binary Tree

        Left Wing | Right Wing
                  |
Level 0          (M)
                / | \
               /  |  \
Level 1      (N)  |  (O)
             / \  |  /
Level 2    (P) (Q)|(R)
                  |


(M) has total of 3 nodes in the left wing and total of 2 nodes in the right wing. (M)'s 3 left
wing are (N), (P) and (Q) while (M)'s right wing are (O) and (R).

A binary tree has levels, like layers in a pyramid. Level 0 is the topmost level. Level 1 is the
first level beneath Level 0, Level 2 is beneath Level 1, and so on.

In the binary mlm system, a member is assigned a node or slot. He is supposed to recruit or
sponsor new members into the system making the binary tree grow downward. A member can only have
no more than two direct downlines, but he can sponsor or be a direct referral to any member.

You will find in the database that a member has many direct referrals (DR). Direct Referral is
different from the Upline Code. The DirectReferral field is the code of the member who
sponsored the new member. In the above example: (M) recruited (N) and (O). (M)'s member code is
'000001', (N) is '000002', (O) is '000003'. In (N)'s record, his DirectReferral field will be
'000001' and UplineCode is '000001' because (N) recruited and at the same time a direct downline
of (M). In (O)'s record, his DirectReferral is '000001' and UplineCode is '000002'. (O) is
sponsored by (M) but his upline is (N).

The reason for multiple sponsorships or referrals is that there is a commission for every
referral a member makes. This is on top of the commission he gets when he makes a pair in his
left and right wing. I'm not gonna discuss about the details of the computation of commissions
because its another subject, but if you're familiar with binary mlm you'll find this module
useful in computing the commissions in this type of mktg strategy. By modifying the GetGenealogy
procedure in the form code and passing the DateJoined and MemberStatus fields, you can make
computations and generate commission reports.

The program generates the genealogy report using the member's code and the upline code.
I converted the records in the database to text file and then from text file, i processed
the genealogy report and returned it as array elements. Actually there are many ways to process
the report and there are faster processes in generating it but its up to you to study it. The
program is not commented. If you're interested in discussing the other details of the
program, you can email me or post your questions.

Sorry if i can't write well in english, but i tried my best. This is probably my last
submission so i hope you all like it. Im gonna start my new job as a materials keeper in a
printing business so i have little time making new programs.

BTW, feel free to modify and use this program. I can say that the program is accurate in
giving the genealogy of a member but if you're gonna use it AS IS for business or commercial,
you take the risk. This program is designed for beginners and for studying use only.

Some modules i used in the program came from other sources. I will mention them in the message
posts next time. If you will also notice, i used the names of the authors in NWind to fill up
the names in my database, hope they don't mind. Its just for temporary data. Whtevr...

This program is dedicated to Ms. Sheryl L. Taripe. 2 (,") (",) 28