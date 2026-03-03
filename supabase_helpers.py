"""
Supabase 헬퍼 함수 모음 (v2.1 - 하이브리드 검색)
- 테이블: test_cases_v21, spec_docs_v21
- 기능: 벡터 검색 + LLM 재랭킹
"""

import streamlit as st
from supabase import create_client, Client
import google.generativeai as genai
import json
from datetime import datetime
import uuid
import numpy as np

# ========================================
# 환경 변수 로드
# ========================================
SUPABASE_URL = st.secrets.get("SUPABASE_URL")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY")
GOOGLE_API_KEY = st.secrets.get("GOOGLE_API_KEY")
TABLE_NAME = st.secrets.get("TABLE_NAME", "test_cases_v21")
SPEC_TABLE_NAME = st.secrets.get("SPEC_TABLE_NAME", "spec_docs_v21")

# 하이브리드 검색 설정
INITIAL_SEARCH_COUNT = st.secrets.get("INITIAL_SEARCH_COUNT", 30)
FINAL_SEARCH_COUNT = st.secrets.get("FINAL_SEARCH_COUNT", 10)
RERANK_METHOD = st.secrets.get("RERANK_METHOD", "gemini")

# Gemini 설정
genai.configure(api_key=GOOGLE_API_KEY)


# ========================================
# Supabase 클라이언트
# ========================================
def get_supabase_client() -> Client:
    """Supabase 클라이언트 반환"""
    try:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        st.error(f"❌ Supabase 연결 실패: {str(e)}")
        return None


# ========================================
# 임베딩 생성
# ========================================
def generate_embedding(text: str):
    """텍스트를 768차원 벡터로 변환 (Gemini gemini-embedding-001)"""
    try:
        result = genai.embed_content(
            model="models/gemini-embedding-001",
            content=text,
            task_type="retrieval_document",
            output_dimensionality=768
        )
        return result['embedding']
    except Exception as e:
        st.error(f"❌ 임베딩 생성 실패: {str(e)}")
        return None


# ========================================
# ⭐ 하이브리드 검색 (핵심 기능)
# ========================================
def hybrid_search_test_cases(query_text: str, category_filter=None, limit=None, similarity_threshold=0.3):
    """
    하이브리드 검색: 벡터 검색 → LLM 재랭킹
    
    Args:
        query_text: 사용자 질문
        category_filter: 카테고리 필터 (옵션)
        limit: 검색 개수 제한 (옵션)
        similarity_threshold: 유사도 임계값 (기본: 0.3)
    
    Returns:
        재랭킹된 테스트 케이스 리스트
    """
    supabase = get_supabase_client()
    if not supabase:
        return []
    
    try:
        # limit 파라미터 처리
        if limit:
            initial_count = limit
            final_count = min(limit, FINAL_SEARCH_COUNT)
        else:
            initial_count = INITIAL_SEARCH_COUNT
            final_count = FINAL_SEARCH_COUNT
            
        # 1단계: 벡터 검색 (넓게 가져오기)
        st.info(f"🔍 1단계: 벡터 검색 중... (최대 {INITIAL_SEARCH_COUNT}개)")
        
        query_embedding = generate_embedding(query_text)
        if not query_embedding:
            return []
        
        result = supabase.rpc(
            'match_test_cases_v21',
            {
                'query_embedding': query_embedding,
                'match_count': initial_count,  # limit 적용
                'similarity_threshold': similarity_threshold  # 파라미터 적용
            }
        ).execute()
        
        if not result.data:
            st.warning("⚠️ 벡터 검색 결과가 없습니다.")
            return []
        
        candidates = result.data
        st.success(f"✅ 1단계 완료: {len(candidates)}개 발견")
        
        # 카테고리 필터링
        if category_filter and category_filter != "전체":
            candidates = [c for c in candidates if c.get('category') == category_filter]
            st.info(f"🔖 카테고리 필터 적용: {len(candidates)}개 남음")
        
        # 2단계: LLM 재랭킹
        # st.info(f"🤖 2단계: {RERANK_METHOD.upper()} 재랭킹 중... (상위 {FINAL_SEARCH_COUNT}개 선택)")
        st.info(f"🤖 2단계: {RERANK_METHOD.upper()} 재랭킹 중... (상위 {final_count}개 선택)")
        # reranked = rerank_candidates(query_text, candidates, FINAL_SEARCH_COUNT)
        reranked = rerank_candidates(query_text, candidates, final_count)
        
        st.success(f"✅ 2단계 완료: 최종 {len(reranked)}개 반환")
        
        return reranked
        
    except Exception as e:
        st.error(f"❌ 하이브리드 검색 오류: {str(e)}")
        return []


def hybrid_search_spec_docs(query_text: str, limit=None, similarity_threshold=0.3):
    """
    기획 문서 하이브리드 검색

    Args:
        query_text: 사용자 질문
        limit: 검색 개수 제한 (옵션)
        similarity_threshold: 유사도 임계값 (기본: 0.3)
    """
    supabase = get_supabase_client()
    if not supabase:
        return []
    
    try:
        # limit 처리
        if limit:
            initial_count = limit
            final_count = min(limit // 2, 5)
        else:
            initial_count = 20
            final_count = 5
            
        # 1단계: 벡터 검색
        query_embedding = generate_embedding(query_text)
        if not query_embedding:
            return []
        
        result = supabase.rpc(
            'match_spec_docs_v21',
            {
                'query_embedding': query_embedding,
                'match_count': initial_count,  # limit 적용
                'similarity_threshold': similarity_threshold  # 파라미터 적용
            }
        ).execute()
        
        if not result.data:
            return []
        
        # 2단계: 재랭킹
        # reranked = rerank_candidates(query_text, result.data, 5)  # 상위 5개
        reranked = rerank_candidates(query_text, result.data, final_count)
        
        return reranked
        
    except Exception as e:
        st.error(f"❌ 기획 문서 검색 오류: {str(e)}")
        return []


# ========================================
# ⭐ 재랭킹 로직
# ========================================
def rerank_candidates(query: str, candidates: list, top_k: int):
    """
    후보군을 재랭킹하여 상위 k개 반환
    """
    method = RERANK_METHOD
    
    if method == "gemini":
        return rerank_with_gemini(query, candidates, top_k)
    elif method == "cosine":
        return rerank_with_cosine(query, candidates, top_k)
    elif method == "hybrid":
        return rerank_hybrid(query, candidates, top_k)
    else:
        # 기본: 벡터 검색 결과 그대로
        return candidates[:top_k]


def rerank_with_gemini(query: str, candidates: list, top_k: int):
    """
    Gemini AI를 사용한 관련성 스코어링
    
    각 후보에 대해 0~10점 관련성 점수를 매김
    """
    model = genai.GenerativeModel('gemini-2.0-flash-exp')
    
    scored_candidates = []
    
    progress_bar = st.progress(0)
    total = len(candidates)
    
    for idx, candidate in enumerate(candidates):
        try:
            # 후보 문서 정보 추출
            description = candidate.get('description', '')
            name = candidate.get('name', '')
            category = candidate.get('category', '')
            
            # 데이터에서 추가 정보 추출
            data = candidate.get('data', {})
            if isinstance(data, dict):
                content = data.get('content', '')
                step = data.get('step', '')
                pre_condition = data.get('pre_condition', '')
            else:
                content = ''
                step = ''
                pre_condition = ''
            
            # 텍스트 조합 (최대 500자)
            doc_text = f"""
카테고리: {category}
제목: {name}
설명: {description[:200]}
사전조건: {pre_condition[:100]}
테스트 단계: {step[:100]}
추가내용: {content[:100]}
            """.strip()
            
            # Gemini에게 관련성 평가 요청
            prompt = f"""
당신은 테스트 케이스 관련성 평가 전문가입니다.

[사용자 질문]
{query}

[테스트 케이스]
{doc_text}

위 테스트 케이스가 사용자 질문과 얼마나 관련이 있는지 0~10점으로 평가하세요.

평가 기준:
- 10점: 질문에 직접적으로 답변할 수 있는 완벽한 케이스
- 7~9점: 질문과 매우 관련 있는 케이스
- 4~6점: 질문과 부분적으로 관련 있는 케이스
- 1~3점: 질문과 약간 관련 있는 케이스
- 0점: 전혀 관련 없는 케이스

**반드시 숫자만 출력하세요.** (예: 8)
"""
            
            response = model.generate_content(
                prompt,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.1,
                    max_output_tokens=10
                )
            )
            
            # 점수 추출
            score_text = response.text.strip()
            try:
                score = float(score_text)
            except:
                # 숫자 추출 시도
                import re
                numbers = re.findall(r'\d+\.?\d*', score_text)
                score = float(numbers[0]) if numbers else 5.0
            
            # 점수 범위 제한
            score = max(0, min(10, score))
            
            scored_candidates.append({
                'data': candidate,
                'score': score,
                'vector_similarity': candidate.get('similarity', 0)
            })
            
            # 진행률 업데이트
            progress_bar.progress((idx + 1) / total)
            
        except Exception as e:
            # 에러 발생 시 기본 점수
            scored_candidates.append({
                'data': candidate,
                'score': 5.0,
                'vector_similarity': candidate.get('similarity', 0)
            })
    
    progress_bar.empty()
    
    # 점수 기준 정렬
    scored_candidates.sort(key=lambda x: x['score'], reverse=True)
    
    # 상위 k개 반환 (원본 데이터만)
    return [c['data'] for c in scored_candidates[:top_k]]


def rerank_with_cosine(query: str, candidates: list, top_k: int):
    """
    코사인 유사도 재계산 (정밀)
    
    Supabase 벡터 검색은 근사치이므로, 
    상위 후보들에 대해 정확한 코사인 유사도를 다시 계산
    """
    query_embedding = generate_embedding(query)
    if not query_embedding:
        return candidates[:top_k]
    
    query_vec = np.array(query_embedding)
    scored_candidates = []
    
    for candidate in candidates:
        try:
            # 후보의 임베딩 가져오기 (Supabase에서 반환 안 됨)
            # 대신 description으로 임베딩 재생성
            description = candidate.get('description', '')
            if not description:
                continue
            
            candidate_embedding = generate_embedding(description)
            if not candidate_embedding:
                continue
            
            candidate_vec = np.array(candidate_embedding)
            
            # 코사인 유사도 계산
            cosine_sim = np.dot(query_vec, candidate_vec) / (
                np.linalg.norm(query_vec) * np.linalg.norm(candidate_vec)
            )
            
            scored_candidates.append({
                'data': candidate,
                'score': cosine_sim
            })
            
        except Exception as e:
            scored_candidates.append({
                'data': candidate,
                'score': 0.5
            })
    
    # 유사도 기준 정렬
    scored_candidates.sort(key=lambda x: x['score'], reverse=True)
    
    return [c['data'] for c in scored_candidates[:top_k]]


def rerank_hybrid(query: str, candidates: list, top_k: int):
    """
    하이브리드 재랭킹: Gemini 점수 + 벡터 유사도 혼합
    
    최종 점수 = (Gemini 점수 * 0.7) + (벡터 유사도 * 10 * 0.3)
    """
    model = genai.GenerativeModel('gemini-2.0-flash-exp')
    
    scored_candidates = []
    progress_bar = st.progress(0)
    total = len(candidates)
    
    for idx, candidate in enumerate(candidates):
        try:
            # Gemini 점수 계산 (간소화된 버전)
            description = candidate.get('description', '')[:300]
            name = candidate.get('name', '')
            
            prompt = f"""
질문: {query}
테스트: {name} - {description}

관련성 점수 (0~10): """
            
            response = model.generate_content(
                prompt,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.1,
                    max_output_tokens=5
                )
            )
            
            import re
            numbers = re.findall(r'\d+\.?\d*', response.text.strip())
            gemini_score = float(numbers[0]) if numbers else 5.0
            gemini_score = max(0, min(10, gemini_score))
            
            # 벡터 유사도 (0~1 → 0~10 스케일)
            vector_score = candidate.get('similarity', 0.5) * 10
            
            # 혼합 점수
            final_score = (gemini_score * 0.7) + (vector_score * 0.3)
            
            scored_candidates.append({
                'data': candidate,
                'score': final_score,
                'gemini_score': gemini_score,
                'vector_score': vector_score
            })
            
            progress_bar.progress((idx + 1) / total)
            
        except Exception as e:
            scored_candidates.append({
                'data': candidate,
                'score': 5.0
            })
    
    progress_bar.empty()
    
    # 점수 기준 정렬
    scored_candidates.sort(key=lambda x: x['score'], reverse=True)
    
    return [c['data'] for c in scored_candidates[:top_k]]


# ========================================
# 테스트 케이스 저장 (2.0과 동일)
# ========================================
def save_test_case_to_supabase(test_case_data):
    """
    테스트 케이스를 Supabase에 저장
    
    Args:
        test_case_data: dict 형태의 테스트 케이스
            - input_type: "table_group", "free_form", "file_upload"
            - category, name, link, description, data 등
    
    Returns:
        저장된 케이스 수
    """
    supabase = get_supabase_client()
    if not supabase:
        return 0
    
    input_type = test_case_data.get("input_type", "unknown")
    saved_count = 0
    
    try:
        if input_type == "table_group":
            # 표 형식: 각 행을 개별 케이스로 저장
            group_id = test_case_data.get("group_id")
            if not group_id:
                group_id = f"table_group_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            table_data = test_case_data.get("table_data", [])
            category = test_case_data.get("category", "미분류")
            
            for idx, row in enumerate(table_data, 1):
                # 빈 행 필터링
                if not row.get('CATEGORY') and not row.get('DEPTH 1'):
                    continue
                
                embedding = generate_embedding(
                    f"{row.get('CATEGORY', '')} {row.get('DEPTH 1', '')} "
                    f"{row.get('DEPTH 2', '')} {row.get('STEP', '')}"
                )
                
                insert_data = {
                    "category": category,
                    "name": f"{row.get('DEPTH 1', '')} - {row.get('DEPTH 2', '')}",
                    "link": "",
                    "description": row.get('STEP', ''),
                    "data": {
                        "group_id": group_id,
                        "input_type": "table_group",
                        "no": row.get('NO', idx),
                        "category": row.get('CATEGORY', ''),
                        "depth1": row.get('DEPTH 1', ''),
                        "depth2": row.get('DEPTH 2', ''),
                        "depth3": row.get('DEPTH 3', ''),
                        "pre_condition": row.get('PRE-CONDITION', ''),
                        "step": row.get('STEP', ''),
                        "expect_result": row.get('EXPECT RESULT', '')
                    },
                    "embedding": embedding
                }
                
                supabase.table(TABLE_NAME).insert(insert_data).execute()
                saved_count += 1
        
        elif input_type == "free_form":
            # 줄글 형식: 단일 케이스로 저장
            embedding = generate_embedding(
                f"{test_case_data.get('name', '')} {test_case_data.get('description', '')}"
            )
            
            insert_data = {
                "category": test_case_data.get("category", "미분류"),
                "name": test_case_data.get("name", ""),
                "link": test_case_data.get("link", ""),
                "description": test_case_data.get("description", ""),
                "data": {
                    "input_type": "free_form",
                    "content": test_case_data.get("content", "")
                },
                "embedding": embedding
            }
            
            supabase.table(TABLE_NAME).insert(insert_data).execute()
            saved_count = 1
        
        elif input_type == "file_upload":
            # 파일 업로드: 각 행을 개별 케이스로 저장
            file_data = test_case_data.get("file_data", [])
            category = test_case_data.get("category", "미분류")
            
            for row in file_data:
                if not row.get('제목'):
                    continue
                
                embedding = generate_embedding(
                    f"{row.get('제목', '')} {row.get('내용', '')}"
                )
                
                insert_data = {
                    "category": category,
                    "name": row.get('제목', ''),
                    "link": row.get('링크', ''),
                    "description": row.get('내용', ''),
                    "data": {
                        "input_type": "file_upload",
                        "content": row.get('추가정보', '')
                    },
                    "embedding": embedding
                }
                
                supabase.table(TABLE_NAME).insert(insert_data).execute()
                saved_count += 1
        
        return saved_count
        
    except Exception as e:
        st.error(f"❌ 저장 실패: {str(e)}")
        return 0


def save_spec_doc_to_supabase(spec_doc_data):
    """
    기획 문서를 Supabase에 저장
    """
    supabase = get_supabase_client()
    if not supabase:
        return False
    
    try:
        embedding = generate_embedding(
            f"{spec_doc_data.get('title', '')} {spec_doc_data.get('content', '')}"
        )
        
        insert_data = {
            "title": spec_doc_data.get("title", ""),
            "doc_type": spec_doc_data.get("doc_type", "Notion"),
            "link": spec_doc_data.get("link", ""),
            "content": spec_doc_data.get("content", ""),
            "embedding": embedding
        }
        
        supabase.table(SPEC_TABLE_NAME).insert(insert_data).execute()
        return True
        
    except Exception as e:
        st.error(f"❌ 기획 문서 저장 실패: {str(e)}")
        return False
