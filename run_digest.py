#!/usr/bin/env python3
"""Weekly AI-for-Polymers digest generator — 2026-06-22"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

DATE = "2026-06-22"

data = {
    "date": DATE,
    "headline": (
        "Foundational generative models hit the lab bench, autonomous multi-agent "
        "polymer informatics platforms go live, and AI recycling systems claim 60 pct "
        "efficiency gains as the field shifts from tools to end-to-end pipelines."
    ),
    "executive_summary": (
        "The week of 22 June 2026 marks an inflection point in AI-driven polymer "
        "science: generative foundation models are no longer just predicting properties "
        "but designing wholly novel structures that pass laboratory validation. Georgia "
        "Tech's POLYT5 encoder-decoder model produced a new dielectric polymer that "
        "passed physical experiments, representing the first foundational model for "
        "generative polymer design with experimental confirmation. Simultaneously, "
        "IBM's WO 2026 patent filing describes an inverse-design paradigm in which "
        "foundation models receive natural-language property specifications and return "
        "chemically valid candidate structures without human bottlenecks.\n\n"
        "On the data infrastructure front, the Open Polymers 2026 (OPoly26) dataset "
        "has become the most comprehensive DFT-grounded polymer resource in existence, "
        "covering more than 6.35 million calculations across 360-atom clusters and "
        "over 1.2 billion atoms. The companion PolyMon unified framework (arXiv "
        "2603.13303) and the autonomous multi-agent system described in arXiv "
        "2602.00103 now allow property prediction, generative design, and bio-polymer "
        "discovery to be chained into a single workflow with minimal human "
        "intervention. OPoly26 captures monomer composition diversity, degree of "
        "polymerisation, chain architecture variations, and solvation environments "
        "in a single unified benchmarking resource.\n\n"
        "Characterization is being transformed by deep-learning models for FTIR "
        "fingerprint-region prediction, with polystyrene reaching 87 percent accuracy, "
        "HDPE 84 percent, and PET 81 percent in a 2026 MRS Advances study. A parallel "
        "PMC review highlights AI-aided crystallization elution fractionation (CEF) "
        "for automated polyolefin resin assessment, and symbiotic human-AI workflows "
        "are becoming the new standard in polymer spectroscopy labs. Inline NIR and "
        "Raman tools now feed machine-learning feedback loops that adjust synthesis "
        "parameters in real time.\n\n"
        "Sustainability and recycling are the most commercially urgent application "
        "area. AI-driven sorting systems combining hyperspectral imaging and "
        "deep-learning identification are achieving 60 percent efficiency gains in "
        "recycling throughput. Multi-objective AI optimisation of plastic waste "
        "pyrolysis is simultaneously minimising energy input and maximising carbon "
        "return on investment. The AI Circular Economy Conference 2026 placed "
        "depolymerisation process control, digital product passports, and "
        "biodegradability prediction models front and centre, reflecting industry's "
        "shift from sustainability as a niche goal to circularity as a core "
        "competitive strategy."
    ),
    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5 Foundation Model Validates First Lab-Tested Generative Polymer Design",
            "description": (
                "Georgia Tech researchers unveiled POLYT5, an encoder-decoder foundation "
                "chemical language model trained on experimentally verified polymers that "
                "accepts desired property specifications and outputs novel polymer "
                "structures. The model produced a new dielectric polymer that passed "
                "laboratory validation — the first time a generative polymer foundation "
                "model's output has been physically confirmed. The approach couples "
                "AI-driven candidate generation with automated flow chemistry and "
                "real-time analytical feedback, enabling experimental validation without "
                "human bottlenecks."
            ),
            "source": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "OPoly26 Dataset: 6.35 Million DFT Calculations Reshape Polymer Benchmarking",
            "description": (
                "The Open Polymers 2026 (OPoly26) dataset provides more than 6.35 million "
                "density functional theory calculations derived from polymeric systems "
                "encompassing over 1.2 billion total atoms across clusters of up to 360 "
                "atoms. OPoly26 captures monomer composition diversity, degree of "
                "polymerisation, chain architectures, and solvation environments in a "
                "single unified resource. The dataset is already enabling hybrid "
                "simulation-ML frameworks for thermal degradation prediction in "
                "high-performance polyimides and adaptive GNN models for multitask "
                "thermal property prediction."
            ),
            "source": "https://arxiv.org/abs/2512.23117",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "PolyMon Unified Framework Bridges Multi-Task Polymer Property Prediction",
            "description": (
                "PolyMon (arXiv 2603.13303) introduces a unified framework for polymer "
                "property prediction that fuses graph neural network fingerprints with "
                "adaptive molecular representations in a multitask learning architecture. "
                "The system targets thermal properties including glass transition "
                "temperature, melting point, and decomposition temperature simultaneously, "
                "reducing the need for separate per-property models. Benchmarked against "
                "OPoly26, PolyMon demonstrates improved generalisation across sparse data "
                "regions that challenge conventional single-task regressors."
            ),
            "source": "https://arxiv.org/abs/2603.13303",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "Autonomous Multi-Agent AI System Integrates Full Polymer Informatics Pipeline",
            "description": (
                "A new autonomous multi-agent AI system (arXiv 2602.00103) spans the "
                "complete polymer informatics stack: property prediction, generative "
                "design, and bio-polymer discovery for both synthetic and natural polymer "
                "families. Agent modules communicate through structured tool-use APIs, "
                "enabling high-throughput screening campaigns to run without continuous "
                "human supervision. The architecture incorporates polymer-specific domain "
                "knowledge and the OPoly26 dataset as a shared knowledge base."
            ),
            "source": "https://arxiv.org/abs/2602.00103",
            "impact": "high",
        },
        {
            "category": "Characterization",
            "title": "Deep Learning Pushes FTIR Fingerprint-Region Prediction Past 87 Pct Accuracy",
            "description": (
                "A 2026 MRS Advances study optimised deep-learning models for FTIR "
                "spectral prediction in the polymer fingerprint region, achieving 87 "
                "percent accuracy for polystyrene, 84 percent for HDPE, and 81 percent "
                "for PET — substantially exceeding prior reference-library methods. "
                "Inline NIR and Raman tools now feed similar ML feedback loops that "
                "adjust synthesis parameters in real time during polymerisation. A "
                "parallel PMC review highlights AI-aided CEF for automated polyolefin "
                "resin assessment as an emerging quality-control standard."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "high",
        },
        {
            "category": "Processing",
            "title": "AI Manufacturing Cuts Cycle Times 30 Pct and Off-Spec Production by 2 Pct",
            "description": (
                "A 2026 MRS Communications review of AI in advanced polymer manufacturing "
                "documents that physics-informed neural networks predicting optimal heating "
                "curves reduce injection moulding cycle times by up to 30 percent, while "
                "real-time ML process control cuts off-spec production by more than 2 "
                "percent and reduces natural-gas energy consumption by 10-20 percent. "
                "Digital twins simulating entire production chains enable multi-process "
                "optimisation, and predictive quality systems now exceed 99 percent "
                "accuracy in defect detection."
            ),
            "source": "https://link.springer.com/article/10.1557/s43579-026-00972-5",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "AI Recycling Systems Report 60 Pct Efficiency Gains; Pyrolysis Optimisation Scales",
            "description": (
                "AI-driven sorting systems combining hyperspectral imaging and deep-learning "
                "polymer identification are achieving up to 60 percent efficiency gains in "
                "recycling plant throughput according to 2026 industry data. Multi-objective "
                "AI optimisation of plastic waste pyrolysis (Polymers 18(9):1062) jointly "
                "minimises energy input and maximises carbon return on investment, making "
                "thermochemical recycling economically viable at scale. ML models for "
                "biodegradability prediction are enabling upstream design of polymers with "
                "built-in end-of-life pathways."
            ),
            "source": "https://doi.org/10.3390/polym18091062",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "LLMs Benchmarked for Polymer Property Prediction; Solubility Validated in ACS",
            "description": (
                "A PubMed-indexed benchmarking study evaluated fine-tuned LLMs against "
                "traditional ML methods for polymer property classification, finding that "
                "LLM-based methods match or exceed conventional approaches without requiring "
                "complex molecular fingerprinting or large labelled datasets. ACS Materials "
                "Letters validated LLM-based polymer solubility prediction using "
                "natural-language property descriptions as direct model inputs. Advanced "
                "Functional Materials (2026) confirmed that LLMs are accelerating synthesis "
                "planning, literature mining, and retrosynthesis across materials science."
            ),
            "source": "https://pubmed.ncbi.nlm.nih.gov/41082689/",
            "impact": "medium",
        },
        {
            "category": "LLMs",
            "title": "Agentic Intelligence Paradigm Emerges for Autonomous Materials Discovery",
            "description": (
                "A 2026 arXiv perspective (2602.00169) formalises agentic intelligence "
                "for materials science — LLM-orchestrated systems that plan multi-step "
                "experimental campaigns, invoke computational tools, retrieve literature, "
                "and iterate without human input between cycles. Polymer informatics is "
                "identified as the leading testbed due to the availability of large "
                "structured datasets (OPoly26, PI1M) and established property prediction "
                "benchmarks. The paper argues self-driving polymer laboratories will be "
                "operational within 18-24 months as agent frameworks mature."
            ),
            "source": "https://arxiv.org/abs/2602.00169",
            "impact": "medium",
        },
        {
            "category": "Generative Design",
            "title": "IBM Patent Filing Covers Foundation-Model-Driven Substructural Polymer Replacement",
            "description": (
                "IBM's WO 2026 patent filing introduces an inverse-design paradigm where "
                "a foundation model receives user prompts specifying desired properties "
                "and generates chemically valid substructural replacements predicted to "
                "exhibit those properties — bypassing traditional forward-screening "
                "workflows. The filing describes integration with automated synthesis "
                "hardware so AI-generated candidates can be synthesised and characterised "
                "in a closed loop. This industrial-scale approach complements academic "
                "foundation models like POLYT5 and signals significant IP activity in "
                "generative polymer design."
            ),
            "source": "https://www.patsnap.com/resources/blog/rd-blog/ai-generative-polymer-design-2026-patsnap-eureka/",
            "impact": "medium",
        },
    ],
    "key_trends": [
        "Generative foundation models move from academic benchmarks to lab-validated polymer synthesis",
        "Autonomous multi-agent pipelines collapse the property prediction-design-synthesis cycle",
        "AI recycling and sustainability applications become core industrial strategy, not niche R&D",
        "LLMs replace classical feature engineering for polymer informatics classification tasks",
        "Mega-datasets (OPoly26, 6.35M DFT calculations) democratise high-accuracy ML training",
    ],
    "notable_tools": [
        {
            "name": "POLYT5",
            "purpose": "Encoder-decoder foundation language model for generative inverse polymer design; accepts property specs and outputs lab-validated novel structures.",
            "url": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
        },
        {
            "name": "OPoly26 Dataset",
            "purpose": "Open benchmark dataset of 6.35M DFT calculations across diverse polymer systems covering composition, architecture, and solvation.",
            "url": "https://arxiv.org/abs/2512.23117",
        },
        {
            "name": "PolyMon",
            "purpose": "Unified multi-task polymer property prediction framework fusing GNN fingerprints with adaptive representations for thermal properties.",
            "url": "https://arxiv.org/abs/2603.13303",
        },
        {
            "name": "Polymer Genome",
            "purpose": "Data-powered polymer informatics platform for property predictions and AI-assisted retrosynthesis planning at polymergenome.org.",
            "url": "https://polymergenome.org",
        },
        {
            "name": "PI1M",
            "purpose": "Benchmark database of approximately 1 million virtual polymers generated by a generative model trained on PolyInfo experimental data.",
            "url": "https://pubs.acs.org/doi/abs/10.1021/acs.jcim.0c00726",
        },
        {
            "name": "Imubit Polymer Process AI",
            "purpose": "Industrial AI platform for real-time polymer processing optimisation, delivering 99%+ defect detection accuracy and 10-20% energy reduction.",
            "url": "https://imubit.com/articles/polymer-processing",
        },
    ],
    "outlook": (
        "The convergence of generative foundation models, autonomous agent pipelines, "
        "and mega-scale DFT datasets is compressing the polymer discovery cycle from "
        "years to weeks. Within the next 12-18 months, self-driving polymer laboratories "
        "that autonomously iterate from property specification through synthesis and "
        "characterisation are expected to transition from prototypes to operational "
        "facilities at leading industrial and academic centres. Sustainability will be "
        "the primary commercial driver: regulatory pressure on plastic waste and carbon "
        "footprint is accelerating investment in AI-optimised recycling, "
        "biodegradability-by-design, and closed-loop circular manufacturing. LLMs will "
        "increasingly serve as the natural-language interface layer making these "
        "capabilities accessible to non-specialist polymer engineers, while agentic "
        "systems handle the underlying computational complexity. The combination of "
        "IBM-scale IP filings and Georgia Tech-style academic validation signals that "
        "generative polymer AI is graduating from research curiosity to industrially "
        "deployable technology."
    ),
}

if __name__ == "__main__":
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    docx_path = output_dir / f"Polymer_AI_Digest_{DATE}.docx"
    pptx_path = output_dir / f"Polymer_AI_Digest_{DATE}.pptx"

    print(f"Generating Word document -> {docx_path}")
    generate_word(data, docx_path)

    print(f"Generating PowerPoint   -> {pptx_path}")
    generate_pptx(data, pptx_path)

    print("\nDone.")
