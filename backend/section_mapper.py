import logging
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

logger = logging.getLogger(__name__)

class SectionMapper:
    def __init__(self, model_name="all-MiniLM-L6-v2"):
        logger.info(f"Loading SentenceTransformer model: {model_name}")
        self.model = SentenceTransformer(model_name)

    def map_sections(self, input_sections, template_sections):
        if not input_sections or not template_sections:
            return {}
            
        input_titles = [s.title for s in input_sections]
        template_titles = [s.title for s in template_sections]

        input_embeddings = self.model.encode(input_titles)
        template_embeddings = self.model.encode(template_titles)

        similarity_matrix = cosine_similarity(input_embeddings, template_embeddings)

        mapping = {}
        for i, in_sec in enumerate(input_sections):
            if not in_sec.title or in_sec.title == "Document Start":
                continue
            best_match_idx = np.argmax(similarity_matrix[i])
            best_score = similarity_matrix[i][best_match_idx]
            
            mapping[in_sec.title] = template_sections[best_match_idx]
            logger.info(f"Mapped '{in_sec.title}' -> '{template_sections[best_match_idx].title}' (Score: {best_score:.2f})")
            
        return mapping
