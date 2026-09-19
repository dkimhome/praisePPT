def mix_lyrics_files(kor_file, eng_file, output_file):
    # 1. 한글 파일과 영어 파일 읽어오기
    with open(kor_file, 'r', encoding='utf-8') as f:
        kor_lines = f.readlines()
        
    with open(eng_file, 'r', encoding='utf-8') as f:
        eng_lines = f.readlines()
        
    # 2. 빈 줄 지우기 (실수로 들어간 엔터 방지)
    kor_lines = [line.strip() for line in kor_lines if line.strip()]
    eng_lines = [line.strip() for line in eng_lines if line.strip()]
    
    mixed_content = []
    
    # 3. 지퍼(zip) 채우듯이 한글 한 줄, 영어 한 줄 묶어주기
    for kor, eng in zip(kor_lines, eng_lines):
        # 한글 가사 밑에 영어 가사를 붙이고 묶음 저장
        mixed_content.append(f"{kor}\n{eng}")
        
    # 4. 묶음들 사이에 빈 줄(엔터 2번)을 넣어서 슬라이드 구분해주기
    final_text = '\n\n'.join(mixed_content)
    
    # 5. 최종 결과를 lyrics.txt로 저장
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_text)
        
    print(f"성공! '{output_file}' 파일이 예쁘게 섞여서 만들어졌습니다.")

# 스크립트 실행
mix_lyrics_files('korean.txt', 'english.txt', 'lyrics.txt')