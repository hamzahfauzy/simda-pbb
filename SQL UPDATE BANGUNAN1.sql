USE dbPajak 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'UPDATE_BANGUNAN')
BEGIN
    DROP PROC UPDATE_BANGUNAN
END
GO
create proc UPDATE_BANGUNAN	
@xNOP nvarchar(30),@xNo smallint,
@xNOP1 nvarchar(30),@xNo1 smallint,
@xNOP2 nvarchar(30),@xNo2 smallint,
@xLuas bigint,@xNJOP bigint,@xNOP3 nvarchar(30),
@xJPB nvarchar(2),
@xNOP4 nvarchar(30),@xNo4 smallint,
@xNOP5 nvarchar(30),@xNo5 smallint,
@xNOP6 nvarchar(30),@xNo6 smallint,
@xNOP7 nvarchar(30),@xNo7 smallint,
@xNOP8 nvarchar(30),@xNo8 smallint,
@xNOP9 nvarchar(30),@xNo9 smallint,
@xNOP10 nvarchar(30),@xNo10 smallint,
@xNOP11 nvarchar(30),@xNo11 smallint,
@xNOP12 nvarchar(30),@xNo12 smallint,
@xNOP13 nvarchar(30),@xNo13 smallint,
@xNOP14 nvarchar(30),@xNo14 smallint,
@xNOP15 nvarchar(30),@xNo15 smallint,
@xNOP16 nvarchar(30),@xNo16 smallint,
@xNOP17 nvarchar(30),@xNo17 smallint

as
begin
DELETE FROM DAT_OP_BANGUNAN where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP) =  @xNOP  AND (NO_BNG  =@xNo)
DELETE FROM DAT_NILAI_INDIVIDU where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP1  AND NO_BNG=@xNo1)
Delete from DAT_FASILITAS_BANGUNAN WHERE (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP2  AND NO_BNG=@xNo2)
UPDATE DAT_OBJEK_PAJAK SET TOTAL_LUAS_BNG = @xLuas, NJOP_BNG = @xNJOP where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP3)
IF @xJPB='02' 
Begin
	DELETE FROM DAT_JPB2 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP4 AND NO_BNG=@xNO4 
End
else
if @xJPB='03'
Begin
	DELETE FROM DAT_JPB3 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP5 AND NO_BNG=@xNO5
End
else
if @xJPB='04'
Begin
	DELETE FROM DAT_JPB4 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP6 AND NO_BNG=@xNO6
End
else
if @xJPB='05'
Begin
	DELETE FROM DAT_JPB5 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP7 AND NO_BNG=@xNO7
End
else
if @xJPB='06'
Begin
	DELETE FROM DAT_JPB6 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP8 AND NO_BNG=@xNO8
End
else
if @xJPB='07'
Begin
	DELETE FROM DAT_JPB7 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP9 AND NO_BNG=@xNO9
End
else
if @xJPB='08'
Begin
	DELETE FROM DAT_JPB8 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP10 AND NO_BNG=@xNO10
End
else
if @xJPB='09'
Begin
	DELETE FROM DAT_JPB2 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP11 AND NO_BNG=@xNO11
End
else
if @xJPB='12'
Begin
	DELETE FROM DAT_JPB12 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP12 AND NO_BNG=@xNO12
End
else
if @xJPB='13'
Begin
	DELETE FROM DAT_JPB13 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP13 AND NO_BNG=@xNO13
End
else
if @xJPB='14'
Begin
	DELETE FROM DAT_JPB14 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP14 AND NO_BNG=@xNO14
End
else
if @xJPB='15'
Begin
	DELETE FROM DAT_JPB15 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP15 AND NO_BNG=@xNO15
End
else
if @xJPB='16'
Begin
	DELETE FROM DAT_JPB16 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP16 AND NO_BNG=@xNO16
End
else
begin
	DELETE FROM DAT_JPB17 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP17 AND NO_BNG=@xNO17
end
if @@ERROR <> 0
begin
rollback tran
end
else
begin
commit tran
end
end

