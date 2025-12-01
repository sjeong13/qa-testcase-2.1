# supabase_helpers.py
"""
Supabase 연동 헬퍼 함수들
"""

import streamlit as st
from supabase import create_client
import google.generativeai as genai
import os
from datetime import datetime

# =============================================
# 1. 초기화 함수
# =============================================

@st.cache_resource
def get_supabase_client():
    """Supabase 클라이언트 초기화"""
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception as e:
        st.error(f"Supabase 연결 실패: {str(e)}")
        return None

def get_gemini_embedding_client():
    """Gemini 임베딩용 클라이언트"""
    api_key = os.environ.get("GOOGLE_API_KEY") or st.secrets.get("GOOGLE_API_KEY")
    if not api_key:
        st.error("GOOGLE_API_KEY 환경 변수가 설정되지 않았습니다.")
        return None
    genai.configure(api_key=api_key)
    return True

# =============================================
# 2. 임베딩 생성 함수
# =============================================

def generate_embedding(text):
    """
    텍스트를 768차원 벡터로 변환
    
    Args:
        text (str): 임베딩할 텍스트
    
    Returns:
        list: 768차원 벡터 또는 None
    """
    try:
        if not get_gemini_embedding_client():
            return None
        
        result = genai.embed_content(
            model="models/text-embedding-004",
            content=text,
            task_type="retrieval_document"
        )
        return result['embedding']
    
    except Exception as e:
        st.error(f"임베딩 생성 실패: {str(e)}")
        return None

# =============================================
# 3. 테스트 케이스 저장 함수 (개별 저장 방식)
# =============================================

def save_test_case_to_supabase(test_case):
    """
    단일 테스트 케이스를 Supabase에 저장
    (그룹은 자동으로 개별 케이스로 쪼갬!)
    
    Args:
        test_case (dict): 테스트 케이스 데이터
    
    Returns:
        int: 저장된 케이스 수
    """
    try:
        supabase = get_supabase_client()
        if not supabase:
            return 0
        
        saved_count = 0
        
        # ==========================================
        # Case 1: 표 그룹 데이터 (쪼개서 저장!)
        # ==========================================
        if test_case.get('table_data'):
            group_id = test_case.get('group_id', f"group_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            
            for idx, row in enumerate(test_case.get('table_data', []), 1):
                # 1. 검색용 텍스트 생성
                search_text = (
                    f"{row.get('CATEGORY', '')} "
                    f"{row.get('DEPTH 1', '')} "
                    f"{row.get('DEPTH 2', '')} "
                    f"{row.get('DEPTH 3', '')} "
                    f"{row.get('PRE-CONDITION', '')} "
                    f"{row.get('STEP', '')} "
                    f"{row.get('EXPECT RESULT', '')}"
                )
                
                # 2. 임베딩 생성
                embedding = generate_embedding(search_text)
                if not embedding:
                    st.warning(f"임베딩 생성 실패: {row.get('DEPTH 1', 'unknown')}")
                    continue
                
                # 3. 개별 케이스로 저장
                result = supabase.table('test_cases').insert({
                    "category": row.get('CATEGORY', ''),
                    "name": f"{row.get('DEPTH 1', '')} {row.get('DEPTH 2', '')}".strip(),
                    "link": test_case.get('link', ''),
                    "description": f"[STEP] {row.get('STEP', '')} [EXPECT] {row.get('EXPECT RESULT', '')}",
                    "data": {
                        "group_id": group_id,
                        "input_type": test_case.get('input_type', 'table_group'),
                        "no": row.get('NO', idx),
                        "category": row.get('CATEGORY', ''),
                        "depth1": row.get('DEPTH 1', ''),
                        "depth2": row.get('DEPTH 2', ''),
                        "depth3": row.get('DEPTH 3', ''),
                        "pre_condition": row.get('PRE-CONDITION', ''),
                        "step": row.get('STEP', ''),
                        "expect_result": row.get('EXPECT RESULT', '')
                    },
                    "embedding": str(embedding)
                }).execute()
                
                saved_count += 1
            
            return saved_count
        
        # ==========================================
        # Case 2: 줄글 형식 (그대로 저장)
        # ==========================================
        else:
            # 1. 검색용 텍스트
            search_text = (
                f"{test_case.get('category', '')} "
                f"{test_case.get('name', '')} "
                f"{test_case.get('description', '')}"
            )
            
            # 2. 임베딩 생성
            embedding = generate_embedding(search_text)
            if not embedding:
                st.warning(f"임베딩 생성 실패: {test_case.get('name')}")
                return 0
            
            # 3. 저장
            result = supabase.table('test_cases').insert({
                "category": test_case.get('category', ''),
                "name": test_case.get('name', ''),
                "link": test_case.get('link', ''),
                "description": test_case.get('description', ''),
                "data": test_case,  # 전체 데이터
                "embedding": str(embedding)
            }).execute()
            
            return 1
    
    except Exception as e:
        st.error(f"Supabase 저장 실패: {str(e)}")
        return 0

# =============================================
# 4. 테스트 케이스 불러오기 (그룹 재구성 옵션)
# =============================================
def load_test_cases_from_supabase(limit=None, group_by_id=False):
    """
    Supabase에서 테스트 케이스 불러오기
    
    Args:
        limit (int): 최대 개수 (None이면 전체)
        group_by_id (bool): group_id별로 묶을지 여부
    
    Returns:
        list: 테스트 케이스 리스트
    """
    try:
        supabase = get_supabase_client()
        if not supabase:
            return []
        
        # 전체 조회
        query = supabase.table('test_cases').select('id, category, name, link, description, data, created_at').order('id', desc=True)
        
        if limit:
            query = query.limit(limit)
        
        result = query.execute()
        
        if not group_by_id:
            # 그룹화 안 함 (개별로)
            test_cases = []
            for row in result.data:
                tc = row['data']
                tc['id'] = row['id']
                tc['supabase_id'] = row['id']  # Supabase ID 별도 보관
                test_cases.append(tc)
            return test_cases
        
        else:
            # group_id별로 묶기
            groups = {}
            individuals = []
            
            for row in result.data:
                tc = row['data']
                tc['id'] = row['id']
                tc['supabase_id'] = row['id']
                
                group_id = tc.get('group_id')
                if group_id:
                    if group_id not in groups:
                        groups[group_id] = []
                    groups[group_id].append(tc)
                else:
                    individuals.append(tc)
            
            # 그룹을 table_data 형식으로 재구성
            test_cases = []
            for group_id, items in groups.items():
                group_case = {
                    "id": items[0]['id'],
                    "group_id": group_id,
                    "input_type": items[0].get('input_type', 'table_group'),
                    "category": items[0]['category'],
                    "name": f"({'AI 생성' if 'ai_generated' in group_id else '입력'} 그룹 {len(items)}개)",
                    "table_data": [
                        {
                            'NO': item.get('no', ''),
                            'CATEGORY': item.get('category', ''),
                            'DEPTH 1': item.get('depth1', ''),
                            'DEPTH 2': item.get('depth2', ''),
                            'DEPTH 3': item.get('depth3', ''),
                            'PRE-CONDITION': item.get('pre_condition', ''),
                            'STEP': item.get('step', ''),
                            'EXPECT RESULT': item.get('expect_result', '')
                        }
                        for item in items
                    ]
                }
                test_cases.append(group_case)
            
            # 개별 케이스 추가
            test_cases.extend(individuals)
            
            return test_cases
    
    except Exception as e:
        st.error(f"테스트 케이스 불러오기 실패: {str(e)}")
        return []

# =============================================
# 5. 벡터 유사도 검색
# =============================================

def search_similar_test_cases(query, limit=50, similarity_threshold=0.3):
    """
    벡터 유사도 기반 검색
    
    Args:
        query (str): 검색어
        limit (int): 반환할 최대 결과 수. 100 이상은 검토 필요(Gemini API 호출 시간, API 비용 증가. 노이즈 많음)
        similarity_threshold (float): 최소 유사도 (0~1)
    
    Returns:
        list: 유사한 테스트 케이스 리스트 (유사도 포함)
    """
    try:
        supabase = get_supabase_client()
        if not supabase:
            return []
        
        # 1. 검색어 임베딩
        query_embedding = genai.embed_content(
            model="models/text-embedding-004",
            # model="models/text-embedding-3-small",  # 전문가 찾기 임베딩 모델 (1536차원)
            content=query,
            task_type="retrieval_query"  # 검색용!
        )['embedding']
        
        # 2. RPC 함수 호출 (벡터 검색)
        result = supabase.rpc(
            'match_test_cases',
            {
                'query_embedding': query_embedding,
                'match_count': limit,
                'similarity_threshold': similarity_threshold
            }
        ).execute()
        
        # 3. 결과 파싱
        test_cases = []
        for row in result.data:
            # tc = row['data']  # JSONB 데이터
            tc = row['data'].copy() if row.get('data') else {}  # JSONB 데이터 복사
            tc['id'] = row['id']
            tc['category'] = row.get('category', '')
            tc['name'] = row.get('name', '')
            tc['link'] = row.get('link', '')
            tc['description'] = row.get('description', '')
            tc['similarity'] = row['similarity']  # 유사도 추가!
            # tc['similarity'] = row.get('similarity', 0)
            
            test_cases.append(tc)
        
        return test_cases
    
    except Exception as e:
        st.error(f"벡터 검색 실패: {str(e)}")
        return []

# =============================================
# 6. 테스트 케이스 삭제
# =============================================

def delete_test_case_from_supabase(test_case_id):
    """
    테스트 케이스 삭제
    
    Args:
        test_case_id (int): Supabase ID
    
    Returns:
        bool: 성공 여부
    """
    try:
        supabase = get_supabase_client()
        if not supabase:
            return False
        
        supabase.table('test_cases').delete().eq('id', test_case_id).execute()
        return True
    
    except Exception as e:
        st.error(f"삭제 실패: {str(e)}")
        return False

# =============================================
# 7. 기획 문서 함수들 (테스트 케이스와 동일 구조)
# =============================================

def save_spec_doc_to_supabase(spec_doc):
    """기획 문서 저장"""
    try:
        supabase = get_supabase_client()
        if not supabase:
            return False
        
        # 검색용 텍스트
        search_text = f"{spec_doc.get('title', '')} {spec_doc.get('content', '')}"
        
        # 임베딩 생성
        embedding = generate_embedding(search_text)
        if not embedding:
            return False
        
        # 저장
        result = supabase.table('spec_docs').insert({
            "title": spec_doc.get('title', ''),
            "doc_type": spec_doc.get('doc_type', ''),
            "link": spec_doc.get('link', ''),
            "content": spec_doc.get('content', ''),
            "embedding": str(embedding)
        }).execute()
        
        return True
    
    except Exception as e:
        st.error(f"기획 문서 저장 실패: {str(e)}")
        return False

def load_spec_docs_from_supabase():
    """기획 문서 불러오기"""
    try:
        supabase = get_supabase_client()
        if not supabase:
            return []
        
        result = supabase.table('spec_docs').select('*').order('id', desc=True).execute()
        return result.data
    
    except Exception as e:
        st.error(f"기획 문서 불러오기 실패: {str(e)}")
        return []

def search_similar_spec_docs(query, limit=50, similarity_threshold=0.3):
    """기획 문서 벡터 검색"""
    try:
        supabase = get_supabase_client()
        if not supabase:
            return []
        
        # 검색어 임베딩
        query_embedding = genai.embed_content(
            model="models/text-embedding-004",
            content=query,
            task_type="retrieval_query"
        )['embedding']
        
        # 벡터 검색
        result = supabase.rpc(
            'match_spec_docs',
            {
                'query_embedding': query_embedding,
                'match_count': limit,
                'similarity_threshold': similarity_threshold
            }
        ).execute()
        
        return result.data
    
    except Exception as e:
        st.error(f"기획 문서 검색 실패: {str(e)}")
        return []
