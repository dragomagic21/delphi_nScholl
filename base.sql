/******************************************************************************/
/****         Generated by IBExpert 2005.09.25 25.09.2009 11:25:33         ****/
/******************************************************************************/

SET SQL DIALECT 3;

SET NAMES WIN1251;

CREATE DATABASE 'D:\nSchool\BASE.FDB'
USER 'SYSDBA' PASSWORD 'masterkey'
PAGE_SIZE 16384
DEFAULT CHARACTER SET WIN1251;



/******************************************************************************/
/****                               Domains                                ****/
/******************************************************************************/

CREATE DOMAIN D_C2 AS
CHAR(10)
COLLATE WIN1251_UA;

CREATE DOMAIN D_DATE AS
DATE
DEFAULT '21.07.1988';

CREATE DOMAIN D_FLOAT AS
FLOAT;

CREATE DOMAIN D_ID AS
INTEGER
NOT NULL;

CREATE DOMAIN D_INT AS
INTEGER;

CREATE DOMAIN D_INT0 AS
INTEGER
DEFAULT 0
NOT NULL;

CREATE DOMAIN D_LOG AS
BLOB SUB_TYPE 1 SEGMENT SIZE 80;

CREATE DOMAIN D_VC1 AS
CHAR(1)
COLLATE WIN1251_UA;

CREATE DOMAIN D_VC100 AS
VARCHAR(100)
COLLATE WIN1251_UA;

CREATE DOMAIN D_VC15 AS
VARCHAR(15)
COLLATE WIN1251_UA;

CREATE DOMAIN D_VC21 AS
VARCHAR(21)
COLLATE WIN1251_UA;

CREATE DOMAIN D_VC4 AS
VARCHAR(10)
COLLATE WIN1251_UA;

CREATE DOMAIN D_VC6 AS
VARCHAR(6)
DEFAULT '1';

CREATE DOMAIN IMG AS
BLOB SUB_TYPE 0 SEGMENT SIZE 80;



/******************************************************************************/
/****                              Generators                              ****/
/******************************************************************************/

CREATE GENERATOR GEN_TB_AMARKS_ID;
SET GENERATOR GEN_TB_AMARKS_ID TO 52;

CREATE GENERATOR GEN_TB_CLASS_ID;
SET GENERATOR GEN_TB_CLASS_ID TO 17;

CREATE GENERATOR GEN_TB_MARKS_ID;
SET GENERATOR GEN_TB_MARKS_ID TO 49;

CREATE GENERATOR GEN_TB_MEDICINA_ID;
SET GENERATOR GEN_TB_MEDICINA_ID TO 0;

CREATE GENERATOR GEN_TB_PEOP_ID;
SET GENERATOR GEN_TB_PEOP_ID TO 12;

CREATE GENERATOR GEN_TB_PLANCLASS_ID;
SET GENERATOR GEN_TB_PLANCLASS_ID TO 9;

CREATE GENERATOR GEN_TB_PREDMET_ID;
SET GENERATOR GEN_TB_PREDMET_ID TO 4;

CREATE GENERATOR GEN_TB_SCHOOL_ID;
SET GENERATOR GEN_TB_SCHOOL_ID TO 1;

CREATE GENERATOR GEN_TB_TEACHER_ID;
SET GENERATOR GEN_TB_TEACHER_ID TO 3;

CREATE GENERATOR GEN_TB_UPLAN_ID;
SET GENERATOR GEN_TB_UPLAN_ID TO 23;

CREATE GENERATOR GEN_TB_USERS_ID;
SET GENERATOR GEN_TB_USERS_ID TO 3;



/******************************************************************************/
/****                                Tables                                ****/
/******************************************************************************/



CREATE TABLE TB_AMARKS (
    ID       D_ID NOT NULL,
    PEOP     D_ID,
    CLASS    D_ID,
    PREDMET  D_VC15,
    TEACHER  D_VC21,
    O1       D_VC6,
    O2       D_VC6,
    YER      D_VC6
);

CREATE TABLE TB_CLASS (
    ID      D_ID NOT NULL,
    SCHOOL  D_ID,
    NUM     D_ID,
    NAME    D_VC4,
    LOG     D_LOG,
    UPLAN   D_INT
);

CREATE TABLE TB_MARKS (
    ID       D_ID NOT NULL,
    PEOP     D_ID,
    PREDMET  D_ID,
    O1       D_VC6,
    O2       D_VC6,
    YER      D_VC6,
    FL       D_INT
);

CREATE TABLE TB_MEDICINA (
    ID    D_INT0 NOT NULL,
    PEOP  D_ID,
    DB    D_DATE,
    DE    D_DATE,
    TXT   D_VC100
);

CREATE TABLE TB_PEOP (
    ID        D_ID NOT NULL,
    CLASS     D_ID,
    FNAME     D_VC15,
    NAME      D_VC15,
    SNAME     D_VC15,
    BIRTHDAY  D_DATE,
    ADR       D_VC100,
    PS        D_C2,
    PN        D_INT,
    INN       D_VC15,
    PRIM      D_LOG
);

CREATE TABLE TB_PLANCLASS (
    ID       D_ID NOT NULL,
    UPLAN    D_ID,
    TEACHER  D_ID,
    PREDMET  D_ID
);

CREATE TABLE TB_PREDMET (
    ID    D_ID NOT NULL,
    NAME  D_VC15
);

CREATE TABLE TB_SCHOOL (
    ID    D_ID NOT NULL,
    NAME  D_VC15,
    IMG   D_INT
);

CREATE TABLE TB_TEACHER (
    ID        D_ID NOT NULL,
    SCHOOL    D_ID,
    FNAME     D_VC15,
    NAME      D_VC15,
    SNAME     D_VC15,
    BIRTHDAY  D_DATE,
    ADR       D_VC100,
    PS        D_C2,
    PN        D_INT,
    INN       D_VC15,
    PRIM      D_LOG
);

CREATE TABLE TB_UPLAN (
    ID      D_ID NOT NULL,
    NAME    D_VC21,
    SCHOOL  D_ID
);

CREATE TABLE TB_USERS (
    ID     D_ID NOT NULL,
    LOGIN  D_VC4,
    PASS   D_VC4,
    RULES  D_ID
);



/******************************************************************************/
/****                             Primary Keys                             ****/
/******************************************************************************/

ALTER TABLE TB_AMARKS ADD CONSTRAINT PK_TB_AMARKS PRIMARY KEY (ID);
ALTER TABLE TB_CLASS ADD CONSTRAINT PK_TB_CLASS PRIMARY KEY (ID);
ALTER TABLE TB_MARKS ADD CONSTRAINT PK_TB_MARKS PRIMARY KEY (ID);
ALTER TABLE TB_MEDICINA ADD CONSTRAINT PK_TB_MEDICINA PRIMARY KEY (ID);
ALTER TABLE TB_PEOP ADD CONSTRAINT PK_TB_PEOP PRIMARY KEY (ID);
ALTER TABLE TB_PLANCLASS ADD CONSTRAINT PK_TB_PLANCLASS PRIMARY KEY (ID);
ALTER TABLE TB_PREDMET ADD CONSTRAINT PK_TB_PREDMET PRIMARY KEY (ID);
ALTER TABLE TB_SCHOOL ADD CONSTRAINT PK_TB_SCHOOL PRIMARY KEY (ID);
ALTER TABLE TB_TEACHER ADD CONSTRAINT PK_TB_TEACHER PRIMARY KEY (ID);
ALTER TABLE TB_UPLAN ADD CONSTRAINT PK_TB_UPLAN PRIMARY KEY (ID);
ALTER TABLE TB_USERS ADD CONSTRAINT PK_TB_USERS PRIMARY KEY (ID);


/******************************************************************************/
/****                             Foreign Keys                             ****/
/******************************************************************************/

ALTER TABLE TB_AMARKS ADD CONSTRAINT FK_TB_AMARKS_1 FOREIGN KEY (PEOP) REFERENCES TB_PEOP (ID) ON DELETE CASCADE;
ALTER TABLE TB_CLASS ADD CONSTRAINT FK_TB_CLASS_1 FOREIGN KEY (SCHOOL) REFERENCES TB_SCHOOL (ID) ON DELETE CASCADE;
ALTER TABLE TB_MARKS ADD CONSTRAINT FK_TB_MARKS_1 FOREIGN KEY (PEOP) REFERENCES TB_PEOP (ID) ON DELETE CASCADE;
ALTER TABLE TB_MARKS ADD CONSTRAINT FK_TB_MARKS_2 FOREIGN KEY (PREDMET) REFERENCES TB_PREDMET (ID) ON DELETE CASCADE;
ALTER TABLE TB_MEDICINA ADD CONSTRAINT FK_TB_MEDICINA_1 FOREIGN KEY (PEOP) REFERENCES TB_PEOP (ID) ON DELETE CASCADE;
ALTER TABLE TB_PEOP ADD CONSTRAINT FK_TB_PEOP_1 FOREIGN KEY (CLASS) REFERENCES TB_CLASS (ID) ON DELETE CASCADE;
ALTER TABLE TB_PLANCLASS ADD CONSTRAINT FK_TB_PLANCLASS_1 FOREIGN KEY (UPLAN) REFERENCES TB_UPLAN (ID) ON DELETE CASCADE;
ALTER TABLE TB_TEACHER ADD CONSTRAINT FK_TB_TEACHER_1 FOREIGN KEY (SCHOOL) REFERENCES TB_SCHOOL (ID) ON DELETE CASCADE;
ALTER TABLE TB_UPLAN ADD CONSTRAINT FK_TB_UPLAN_1 FOREIGN KEY (SCHOOL) REFERENCES TB_SCHOOL (ID) ON DELETE CASCADE;


/******************************************************************************/
/****                               Indices                                ****/
/******************************************************************************/

CREATE INDEX TB_AMARKS_SORT ON TB_AMARKS (CLASS);
CREATE UNIQUE INDEX TB_CLASS_NAME ON TB_CLASS (NUM, NAME);
CREATE DESCENDING INDEX TB_MED_DB ON TB_MEDICINA (DB);
CREATE INDEX TB_PEOP_IDX1 ON TB_PEOP (FNAME, NAME, SNAME);
CREATE INDEX TB_PLANCLASS_IDX1 ON TB_PLANCLASS (UPLAN);
CREATE INDEX TB_PREDMET_IDX1 ON TB_PREDMET (NAME);
CREATE UNIQUE INDEX TB_SCHOOL_NAME ON TB_SCHOOL (NAME);
CREATE INDEX TB_TEACHER_IDX1 ON TB_TEACHER (FNAME, NAME, SNAME);
CREATE UNIQUE INDEX TB_UPLAN_NAME ON TB_UPLAN (NAME);
CREATE UNIQUE INDEX USER_LOGIN ON TB_USERS (LOGIN, ID);


/******************************************************************************/
/****                               Triggers                               ****/
/******************************************************************************/


SET TERM ^ ;


/******************************************************************************/
/****                         Triggers for tables                          ****/
/******************************************************************************/



/* Trigger: TB_AMARKS_BI */
CREATE TRIGGER TB_AMARKS_BI FOR TB_AMARKS
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_AMARKS_ID,1);
END
^

/* Trigger: TB_CLASS_BI */
CREATE TRIGGER TB_CLASS_BI FOR TB_CLASS
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_CLASS_ID,1);
END
^

/* Trigger: TB_MARKS_BI */
CREATE TRIGGER TB_MARKS_BI FOR TB_MARKS
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_MARKS_ID,1);
END
^

/* Trigger: TB_MEDICINA_BI */
CREATE TRIGGER TB_MEDICINA_BI FOR TB_MEDICINA
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_MEDICINA_ID,1);
END
^

/* Trigger: TB_PEOP_BI */
CREATE TRIGGER TB_PEOP_BI FOR TB_PEOP
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_PEOP_ID,1);
END
^

/* Trigger: TB_PLANCLASS_BI */
CREATE TRIGGER TB_PLANCLASS_BI FOR TB_PLANCLASS
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_PLANCLASS_ID,1);
END
^

/* Trigger: TB_PREDMET_BI */
CREATE TRIGGER TB_PREDMET_BI FOR TB_PREDMET
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_PREDMET_ID,1);
END
^

/* Trigger: TB_SCHOOL_BI */
CREATE TRIGGER TB_SCHOOL_BI FOR TB_SCHOOL
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_SCHOOL_ID,1);
END
^

/* Trigger: TB_TEACHER_BI */
CREATE TRIGGER TB_TEACHER_BI FOR TB_TEACHER
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_TEACHER_ID,1);
END
^

/* Trigger: TB_UPLAN_BI */
CREATE TRIGGER TB_UPLAN_BI FOR TB_UPLAN
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_UPLAN_ID,1);
END
^

/* Trigger: TB_USERS_BI */
CREATE TRIGGER TB_USERS_BI FOR TB_USERS
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = GEN_ID(GEN_TB_USERS_ID,1);
END
^

SET TERM ; ^

