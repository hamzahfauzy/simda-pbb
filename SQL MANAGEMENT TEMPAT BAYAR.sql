-- ================================================
-- Template generated from Template Explorer using:
-- Create Procedure (New Menu).SQL
--
-- Use the Specify Values for Template Parameters 
-- command (Ctrl-Shift-M) to fill in the parameter 
-- values below.
--
-- This block of comments will not be included in
-- the definition of the procedure.
-- ================================================
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
create proc SP_TAMPIL_BANK
as
select * from tempat_bayar
go
create proc SP_TAMPIL_ID(@cNama nvarchar(30),@cAlamat  nVarchar(50),@KODE nvarchar(2))
as
select * from TEMPAT_BAYAR WHERE NM_TP+'-'+ALAMAT_TP+'-'+KD_TP = @cNama+'-'+@cAlamat+'-'+@KODE
GO

create proc SP_INSERT_BANK 
@cKanwil nvarchar(2),@cKPPBB nvarchar(2),@cBank1 nvarchar(2),
@cBank2 nvarchar(2),@cTP nvarchar(2),@cNama nvarchar(30),@cAlamat nvarchar(50),@cRek nvarchar(15)
as
Begin
insert into TEMPAT_BAYAR (KD_KANWIL,KD_KPPBB,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,[KD_TP],[NM_TP],ALAMAT_TP,NO_REK_TP)
values (@cKanwil,@cKPPBB,@cBank1,@cBank2,@cTP,@cNama,@cAlamat,@cRek)
IF @@ERROR <>0 
BEGIN
ROLLBACK TRANSACTION
END
ELSE
BEGIN
COMMIT TRANSACTION
END
END
GO
CREATE PROC SP_UPDATE_BANK
@cKanwil nvarchar(2),@cKPPBB nvarchar(2),@cBank1 nvarchar(2),
@cBank2 nvarchar(2),@cTP nvarchar(2),@cNama nvarchar(30),@cAlamat nvarchar(50),@cRek nvarchar(15),@xID nvarchar(255)
as
Begin
UPDATE TEMPAT_BAYAR SET KD_KANWIL=@cKanwil,KD_KPPBB=@cKPPBB,KD_BANK_TUNGGAL=@cBank1,KD_BANK_PERSEPSI=@cBank2,[KD_TP]=@cTP,[NM_TP]=@cNama ,ALAMAT_TP=@cAlamat,NO_REK_TP=@cRek
	WHERE NM_TP+'-'+ALAMAT_TP+'-'+KD_TP = @xID 
IF @@ERROR <>0 
BEGIN
ROLLBACK TRANSACTION
END
ELSE
BEGIN
COMMIT TRANSACTION
END
END
go

create proc SP_DEL_BANK
@cNama nvarchar(30),@cAlamat  nVarchar(50),@KODE nvarchar(2)
as
begin
delete TEMPAT_BAYAR WHERE NM_TP+'-'+ALAMAT_TP+'-'+KD_TP = @cNama+'-'+@cAlamat+'-'+@KODE
if @@ERROR<>0
begin
rollback transaction
end
else
begin
commit transaction
end
end
go