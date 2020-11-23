--create database PartDB;
--USE PartDB
--create table Parts(
-- Id int PRIMARY KEY NOT NULL IDENTITY(1,1),
-- Description nvarchar(500),
-- PartNumber nvarchar(250),
-- Link nvarchar(250),
-- Price nvarchar(250)
--)

--create table Catergoery(
--Id int PRIMARY KEY NOT NULL Identity,
--Name nvarchar(100),
--Description nvarchar(250)
--)

--ALTER TABLE Parts
--ADD CatergoeryId int FOREIGN KEY REFERENCES Catergoery(Id)

--create table BOMPart(
-- Id int PRIMARY KEY NOT NULL IDENTITY(1,1),
-- JobNumber nvarchar(250) NOT NULL,
-- PartId int NOT NULL FOREIGN KEY REFERENCES Parts(Id),
-- Quantity int NOT NULL,
-- DateCreated datetime NOT NULL 
--)

--ALTER TABLE Parts
--ADD Supplier nvarchar(100)

--insert INTO Catergoery(Name, Description)
--Values('Wire', 'Wire'),
--('Connectors', 'Connectors')

--insert INTO Parts(Description,PartNumber,Link, Price, CatergoeryId, Supplier)
--Values ('Green 16AWG', 'P1234','Fake Link', '$1.25/ft',1,'Good Comapny Store'),
--('Blue 16AWG', 'P134','Fake Link', '$1.35/ft',1,'Good Comapny Store'),
--('Packard 2 Pin Male Connector', 'C134M','Fake Link', '$3.35',2,'We Connector You'),
--('Packard 2 Pin Female Connector', 'C134F','Fake Link', '$3.35',2,'We Connector You')