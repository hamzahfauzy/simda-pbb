
CREATE PROC UPDATE_BUMI
@xProp nVarchar(3), @xKab nVarchar(3), @xxKec nVarchar(3), @xxKel nVarchar(3),
 @xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1), @xNo smallint, 
 @xZNT  nVarchar(2), @Luas bigint , @xJBUMI nVarchar(1) , @xNilai bigint, @xForm nVarchar(11),
 @xStatus int, @xxNOP nvarchar(30),
 @xProp1 nVarchar(3), @xKab1 nVarchar(3), @xxKec1 nVarchar(3), @xxKel1 nVarchar(3),
 @xxBlok1 nVarchar(3), @xxUrut1 nVarchar(4), @xxJenis1 nVarchar(1),
 @xID nvarchar(30),@xForm1 nVarchar(11),@xPersil nvarchar(5),@xJalanOP nvarchar(30),@xBlokOP nvarchar(15),
 @xRW nvarchar(2), @xRT nvarchar(2),@xStatusWP nvarchar(1),@xLBumi bigint, @xNJOP_BM bigint,
 @xTrans nvarchar(1),@xTgl1 datetime, @xNIP1 nvarchar(30),@xTgl2 datetime, @xNIP2 nvarchar(30),
 @xTgl3 datetime, @xNIP3 nvarchar(30),@xCabang smallint,@xLBangunan bigint,@xNJOP_BG bigint,@xPeta smallint,
 @xxNOP1 nvarchar(30)
AS
BEGIN
UPDATE DAT_OP_BUMI SET KD_PROPINSI=@xProp,KD_DATI2=@xKab,KD_KECAMATAN= @xxKec,KD_KELURAHAN=@xxKel,KD_BLOK=@xxBlok,NO_URUT=@xxUrut,KD_JNS_OP=@xxJenis,
			NO_BUMI=@xNo,KD_ZNT=@xZNT,LUAS_BUMI=@Luas,JNS_BUMI=@xJBUMI, NILAI_SISTEM_BUMI=@xNilai,NO_FORMULIR=@xFORM,STATUS_JADI=@xStatus
			WHERE ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)=@xxNOP
UPDATE DAT_OBJEK_PAJAK SET KD_PROPINSI=@xProp1,KD_DATI2=@xKab1,KD_KECAMATAN=@xxKec1,KD_KELURAHAN=@xxKel1,KD_BLOK=@xxBlok1,NO_URUT=@xxUrut1,KD_JNS_OP=@xxJenis1,
			SUBJEK_PAJAK_ID=@xID, NO_FORMULIR_SPOP=@xForm1,NO_PERSIL=@xPersil,JALAN_OP=@xJalanOP, BLOK_KAV_NO_OP=@xBlokOP,RW_OP=@xRW,RT_OP=@xRT,KD_STATUS_WP=@xStatusWP,
			TOTAL_LUAS_BUMI=@xLBumi,NJOP_BUMI=@xNJOP_BM,JNS_TRANSAKSI_OP=@xTrans,TGL_PENDATAAN_OP=@xTgl1,NIP_PENDATA=@xNIP1,TGL_PEMERIKSAAN_OP=@xTgl2,
			NIP_PEMERIKSA_OP=@xNIP2,TGL_PEREKAMAN_OP=@xTgl3,NIP_PEREKAM_OP=@xNIP3,KD_STATUS_CABANG=@xCabang,TOTAL_LUAS_BNG=@xLBangunan,NJOP_BNG=@xNJOP_BG,STATUS_PETA_OP=@xPeta
			WHERE ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)=@xxNOP1
IF @@ERROR<>0 
BEGIN
ROLLBACK TRAN
END
ELSE
BEGIN
COMMIT TRAN
END
END