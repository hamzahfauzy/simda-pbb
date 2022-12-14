USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'HITUNG_DBKB_JPB3')
BEGIN
    DROP PROC HITUNG_DBKB_JPB3
END
GO
create PROCEDURE HITUNG_DBKB_JPB3
@VLC_THN_DBKB_JPB3 NVARCHAR(4)
AS
BEGIN
  DECLARE                           
  @VLN_LBR_BENT_MIN_DBKB_JPB3    INT,
  @VLN_LBR_BENT_MAX_DBKB_JPB3    INT,
  @VLN_TING_KOLOM_MIN_DBKB_JPB3  INT,
  @VLN_TING_KOLOM_MAX_DBKB_JPB3  INT,
  @VLN_HRG_DBKB_JPB3           FLOAT,
  @VLN_NILAI                    BIGINT
  DECLARE C_DBKB_JPB3 CURSOR FOR
  SELECT DISTINCT LBR_BENT_MIN_DBKB_JPB8,LBR_BENT_MAX_DBKB_JPB8,TING_KOLOM_MIN_DBKB_JPB8,TING_KOLOM_MAX_DBKB_JPB8 FROM DBKB_JPB8
  OPEN C_DBKB_JPB3
     FETCH C_DBKB_JPB3 INTO @VLN_LBR_BENT_MIN_DBKB_JPB3,@VLN_LBR_BENT_MAX_DBKB_JPB3,
	      @VLN_TING_KOLOM_MIN_DBKB_JPB3,@VLN_TING_KOLOM_MAX_DBKB_JPB3
     WHILE @@FETCH_STATUS=0
     BEGIN
     SELECT @VLN_HRG_DBKB_JPB3=ROUND((NILAI_DBKB_JPB8*1.3),1) FROM DBKB_JPB8
     WHERE KD_PROPINSI              = '12' AND
	       KD_DATI2                 = '12' AND
           THN_DBKB_JPB8            = @VLC_THN_DBKB_JPB3            AND
           LBR_BENT_MIN_DBKB_JPB8   = @VLN_LBR_BENT_MIN_DBKB_JPB3   AND
           LBR_BENT_MAX_DBKB_JPB8   = @VLN_LBR_BENT_MAX_DBKB_JPB3   AND
           TING_KOLOM_MIN_DBKB_JPB8 = @VLN_TING_KOLOM_MIN_DBKB_JPB3 AND
           TING_KOLOM_MAX_DBKB_JPB8 = @VLN_TING_KOLOM_MAX_DBKB_JPB3

     SELECT @VLN_NILAI=COUNT(*) FROM DBKB_JPB3
     WHERE KD_PROPINSI              = '12' AND
		   KD_DATI2                 = '12' AND
           THN_DBKB_JPB3            = @VLC_THN_DBKB_JPB3            AND
           LBR_BENT_MIN_DBKB_JPB3   = @VLN_LBR_BENT_MIN_DBKB_JPB3   AND
           LBR_BENT_MAX_DBKB_JPB3   = @VLN_LBR_BENT_MAX_DBKB_JPB3   AND
           TING_KOLOM_MIN_DBKB_JPB3 = @VLN_TING_KOLOM_MIN_DBKB_JPB3 AND
           TING_KOLOM_MAX_DBKB_JPB3 = @VLN_TING_KOLOM_MAX_DBKB_JPB3

     IF @VLN_NILAI = 0 
     BEGIN
        INSERT INTO DBKB_JPB3(KD_PROPINSI,KD_DATI2,THN_DBKB_JPB3,LBR_BENT_MIN_DBKB_JPB3,LBR_BENT_MAX_DBKB_JPB3,TING_KOLOM_MIN_DBKB_JPB3,TING_KOLOM_MAX_DBKB_JPB3,NILAI_DBKB_JPB3)
        VALUES ('12','12',@VLC_THN_DBKB_JPB3,@VLN_LBR_BENT_MIN_DBKB_JPB3,@VLN_LBR_BENT_MAX_DBKB_JPB3,
			        @VLN_TING_KOLOM_MIN_DBKB_JPB3,@VLN_TING_KOLOM_MAX_DBKB_JPB3,@VLN_HRG_DBKB_JPB3)
	 END
	 ELSE
	 BEGIN
        UPDATE DBKB_JPB3 SET NILAI_DBKB_JPB3 = @VLN_HRG_DBKB_JPB3
        WHERE KD_PROPINSI              = '12' AND
			  KD_DATI2                 = '12' AND
	          THN_DBKB_JPB3            = @VLC_THN_DBKB_JPB3            AND
	          LBR_BENT_MIN_DBKB_JPB3   = @VLN_LBR_BENT_MIN_DBKB_JPB3   AND
	          LBR_BENT_MAX_DBKB_JPB3   = @VLN_LBR_BENT_MAX_DBKB_JPB3   AND
	          TING_KOLOM_MIN_DBKB_JPB3 = @VLN_TING_KOLOM_MIN_DBKB_JPB3 AND
	          TING_KOLOM_MAX_DBKB_JPB3 = @VLN_TING_KOLOM_MAX_DBKB_JPB3
     END
     FETCH NEXT FROM C_DBKB_JPB3 INTO @VLN_LBR_BENT_MIN_DBKB_JPB3,@VLN_LBR_BENT_MAX_DBKB_JPB3,
	      @VLN_TING_KOLOM_MIN_DBKB_JPB3,@VLN_TING_KOLOM_MAX_DBKB_JPB3
     END
     CLOSE C_DBKB_JPB3
     DEALLOCATE C_DBKB_JPB3
if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
END
