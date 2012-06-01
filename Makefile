# copies most rescent files from eplapp for updating to git.
SERVER=eplapp.library.ualberta.ca
USER=sirsi
REMOTE=/s/sirsi/Unicorn/EPLwork/anisbet/excel.pl
LOCAL=/home/ilsdev/projects/excel/

get:
	scp ${USER}@${SERVER}:${REMOTE} ${LOCAL}

